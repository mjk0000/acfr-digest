"""ACFR Digest — drag-and-drop UI for the audited ACFR General Fund parser.

Wraps parse_cafr (the audited asset) without modifying it: uploads are saved
to a temp directory, run one at a time through parse_cafr.process_pdf, and
written with parse_cafr.write_excel — so the downloaded workbook is identical
to what the CLI produces for the same files.
"""

import csv
import gc
import io
import logging
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from parse_cafr import (
    NOT_FOUND,
    OUTPUT_COLUMNS,
    _NORMALIZABLE_COLS as NUMERIC_COLS,
    ocr_available,
    process_pdf,
    write_excel,
)

try:
    from importlib.metadata import version as _pkg_version
    VERSION = _pkg_version('acfr-digest')
except Exception:
    VERSION = 'dev'

# Fallback-path markers in extraction notes that warrant a manual check.
CAUTION_PATTERN = re.compile(r'verify manually|derived|OCR|inferred', re.IGNORECASE)
GAP_STYLE = 'background-color: #fff3cd; color: #92400e;'

META_COLS = ['Source File', 'Entity Name', 'Fiscal Year End', 'Reporting Units']


def file_notes(row):
    notes = row.get('Extraction Notes')
    if not isinstance(notes, str) or notes == NOT_FOUND:
        return ''
    return notes


def needs_verification(row):
    return bool(CAUTION_PATTERN.search(file_notes(row)))


def _unique(tdir: Path, name: str) -> Path:
    dest, counter = tdir / name, 1
    while dest.exists():
        dest = tdir / f"{Path(name).stem}_{counter}{Path(name).suffix}"
        counter += 1
    return dest


def _stream_to(dest: Path, fobj) -> Path:
    with open(dest, 'wb') as out:
        shutil.copyfileobj(fobj, out, length=1 << 20)
    return dest


def save_uploads(uploads, tdir: Path):
    """Stream uploaded PDFs — and PDFs inside uploaded zips — onto disk.

    Everything is streamed (never .read()/.getvalue() of a whole member):
    hosted capacity is memory-bound, so peak RAM must stay at one buffer,
    not a second full copy of the batch.
    """
    paths = []
    for uf in uploads:
        if uf.name.lower().endswith('.zip'):
            with zipfile.ZipFile(uf) as zf:
                for info in zf.infolist():
                    name = Path(info.filename).name
                    if (info.is_dir() or '__MACOSX' in info.filename
                            or name.startswith('.')
                            or not name.lower().endswith('.pdf')):
                        continue
                    with zf.open(info) as src:
                        paths.append(_stream_to(_unique(tdir, name), src))
        else:
            uf.seek(0)
            paths.append(_stream_to(_unique(tdir, Path(uf.name).name), uf))
    return sorted(paths, key=lambda p: p.name.lower())


def make_logger():
    buf = io.StringIO()
    logger = logging.getLogger('acfr_digest_ui')
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()
    handler = logging.StreamHandler(buf)
    handler.setFormatter(logging.Formatter(
        '%(asctime)s [%(levelname)s] %(message)s', datefmt='%H:%M:%S'))
    logger.addHandler(handler)
    return logger, buf


def checkpoint_csv(results) -> bytes:
    """Same rows the CLI streams to <output>.checkpoint.csv as crash insurance."""
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=OUTPUT_COLUMNS, extrasaction='ignore')
    writer.writeheader()
    for row in results:
        writer.writerow({k: row.get(k, NOT_FOUND) for k in OUTPUT_COLUMNS})
    return buf.getvalue().encode('utf-8')


def run_extraction(workdir: Path, pdf_paths):
    logger, log_buf = make_logger()
    results = []
    n = len(pdf_paths)
    progress = st.progress(0.0, text=f'0 of {n} files done')
    for i, pdf_path in enumerate(pdf_paths):
        with st.status(f'Processing {pdf_path.name} ({i + 1} of {n}) — '
                       'typically 1–3 minutes; scanned/OCR files take '
                       'several more') as status:
            try:
                row = process_pdf(pdf_path, logger)
            except Exception as e:
                row = {col: NOT_FOUND for col in OUTPUT_COLUMNS}
                row['Source File'] = pdf_path.name
                row['Extraction Notes'] = f'Unhandled error: {str(e)[:300]}'
            results.append(row)
            found = sum(1 for c in NUMERIC_COLS if row.get(c, NOT_FOUND) != NOT_FOUND)
            icon = '⚠️' if needs_verification(row) else ('❌' if found == 0 else '✅')
            status.update(state='complete',
                          label=f'{icon} {pdf_path.name} — {found}/{len(NUMERIC_COLS)} '
                                'figures extracted')
        # Free the file's disk and the parser's working set before the next
        # one — long batches must not accumulate.
        pdf_path.unlink(missing_ok=True)
        gc.collect()
        progress.progress((i + 1) / n, text=f'{i + 1} of {n} files done')
    out_path = workdir / 'acfr_digest_results.xlsx'
    write_excel(results, out_path, logger)
    xlsx_bytes = out_path.read_bytes()
    return {'results': results, 'xlsx': xlsx_bytes,
            'csv': checkpoint_csv(results),
            'log': log_buf.getvalue().encode('utf-8')}


def styled_results_table(results):
    # NOT FOUND becomes NaN in the numeric columns so each column keeps a
    # single dtype (Arrow-clean) and gaps are styled/blanked via isna.
    df = pd.DataFrame([
        {'⚠': '⚠️' if needs_verification(r) else '',
         **{c: r.get(c, NOT_FOUND) for c in META_COLS},
         **{c: (None if r.get(c, NOT_FOUND) == NOT_FOUND else r[c])
            for c in NUMERIC_COLS},
         'Extraction Notes': file_notes(r)}
        for r in results
    ])

    def fmt(v):
        if pd.isna(v) or v == NOT_FOUND:
            return ''
        return f'{v:,.2f}' if isinstance(v, float) else v

    def gap(v):
        return GAP_STYLE if pd.isna(v) else ''

    styler = df.style.format(fmt)
    # Styler.map arrived in pandas 2.1; applymap was removed in pandas 3.
    # Attribute access must be lazy — an eager getattr default evaluates
    # the missing attribute and raises.
    map_cells = styler.map if hasattr(styler, 'map') else styler.applymap
    return map_cells(gap, subset=list(NUMERIC_COLS))


def render_results(run):
    results = run['results']
    flagged = [r for r in results if needs_verification(r)]
    gaps = sum(1 for r in results
               for c in NUMERIC_COLS if r.get(c, NOT_FOUND) == NOT_FOUND)

    st.subheader('Results')
    st.markdown(
        'Amber cells are **NOT FOUND** — this parser refuses to guess: a '
        'visible gap beats a plausible wrong number. A blank means "look it '
        'up in the PDF," never zero. Rows marked ⚠️ used a fallback '
        'extraction path (OCR, derived, or inferred figures) — verify those '
        'against the source before use.')
    st.dataframe(styled_results_table(results), width='stretch')

    if flagged:
        with st.expander(f'⚠️ {len(flagged)} file(s) flagged for manual '
                         'verification', expanded=True):
            for r in flagged:
                st.markdown(f"**{r['Source File']}** — {file_notes(r)}")
    if gaps:
        st.caption(f'{gaps} figure(s) came back NOT FOUND across '
                   f'{len(results)} file(s). See the Known limitations tab '
                   'for the PDF classes that cause this.')
    with st.expander('Extraction notes (all files)'):
        for r in results:
            st.markdown(f"**{r['Source File']}**: {file_notes(r) or '_no notes_'}")

    dl1, dl2, dl3 = st.columns(3)
    dl1.download_button(
        '⬇️ Excel workbook (.xlsx)', run['xlsx'],
        file_name='acfr_digest_results.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        type='primary')
    dl2.download_button(
        '⬇️ Checkpoint CSV', run['csv'],
        file_name='acfr_digest_results.checkpoint.csv', mime='text/csv',
        help='The same rows in plain CSV — the CLI writes this file '
             'incrementally as crash insurance.')
    dl3.download_button(
        '⬇️ Run log', run['log'], file_name='acfr_digest_run.log',
        mime='text/plain',
        help='Full parser log for this run — attach it when reporting a '
             'misparse.')


def main():
    st.set_page_config(page_title='ACFR Digest', page_icon='🏛️', layout='wide')
    st.title('🏛️ ACFR Digest')
    st.caption(f'General Fund figures out of Annual Comprehensive Financial '
               f'Reports — alpha {VERSION}')

    tab_extract, tab_limits = st.tabs(['Extract', 'Known limitations'])

    with tab_limits:
        limits = Path(__file__).with_name('KNOWN_LIMITATIONS.md')
        if limits.exists():
            st.markdown(limits.read_text(encoding='utf-8'))
        else:
            st.warning('KNOWN_LIMITATIONS.md not found next to app.py.')

    with tab_extract:
        st.markdown(
            'Upload ACFR PDFs and get back one Excel row per report: Total '
            'Assets, Total Liabilities, the five fund balance categories, '
            'Total Fund Balances, Total Revenues, Total Expenditures, and '
            'Total Other Financing Sources — all from the **General Fund '
            'column only**.')
        if not ocr_available():
            st.info(
                'OCR toolchain not detected (poppler + tesseract + '
                'pytesseract). Native-text PDFs are unaffected, but scanned '
                'or broken-encoding PDFs will return honest NOT FOUNDs '
                'instead of OCR-recovered figures. Install notes are in the '
                'README.')
        # The uploader key is generational: after ingest we bump it and rerun,
        # so Streamlit drops the old widget (and its in-memory upload buffers)
        # before the memory-hungry parsing phase starts.
        gen = st.session_state.setdefault('uploader_gen', 0)
        uploads = st.file_uploader(
            'Drop ACFR PDFs here — or a .zip to upload a whole folder',
            type=['pdf', 'zip'], accept_multiple_files=True,
            key=f'uploads_{gen}',
            help='Browsers cannot drag folders — zip a folder to upload it '
                 'whole.')
        st.caption('Hosted note: one run comfortably handles ~10–15 PDFs '
                   '(a few hundred MB); keep this tab open while it works. '
                   'For bigger batches, use the CLI — see the README.')
        if uploads and st.button('Extract General Fund figures', type='primary'):
            stale = st.session_state.pop('workdir', None)
            if stale:
                shutil.rmtree(stale, ignore_errors=True)
            workdir = Path(tempfile.mkdtemp(prefix='acfr-digest-'))
            pdf_paths = save_uploads(uploads, workdir)
            if not pdf_paths:
                st.error('No PDFs found in the upload.')
                shutil.rmtree(workdir, ignore_errors=True)
            else:
                st.session_state['workdir'] = str(workdir)
                st.session_state['pending'] = [str(p) for p in pdf_paths]
                st.session_state['uploader_gen'] = gen + 1
                st.rerun()
        if 'pending' in st.session_state:
            pdf_paths = [Path(p) for p in st.session_state.pop('pending')]
            workdir = Path(st.session_state['workdir'])
            try:
                st.session_state['run'] = run_extraction(workdir, pdf_paths)
            finally:
                shutil.rmtree(workdir, ignore_errors=True)
                st.session_state.pop('workdir', None)
        if 'run' in st.session_state:
            render_results(st.session_state['run'])


if __name__ == '__main__':
    main()
