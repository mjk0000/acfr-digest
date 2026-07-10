"""acfr-digest console entry point.

Thin wrapper around parse_cafr.main() — the audited parser is wrapped, never
modified. The only addition is a --version flag (sourced from the installed
package metadata), intercepted before delegation; any other invocation is
passed through untouched, so `acfr-digest ...` and `python3 parse_cafr.py ...`
behave identically.
"""

import sys


def package_version() -> str:
    try:
        from importlib.metadata import version
        return version('acfr-digest')
    except Exception:
        return 'dev'


def main():
    if '--version' in sys.argv[1:]:
        print(f"acfr-digest {package_version()}")
        return
    import parse_cafr
    parse_cafr.main()


if __name__ == '__main__':
    main()
