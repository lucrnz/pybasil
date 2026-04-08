"""Command-line interface for pybasil."""

import sys
import argparse
from .interpreter import run


def main():
    """Main entry point for the pybasil CLI."""
    parser = argparse.ArgumentParser(
        prog="pybasil", description="VBScript interpreter in Python"
    )
    parser.add_argument(
        "file",
        nargs="?",
        help="VBScript file to execute (reads from stdin if not specified)",
    )
    parser.add_argument("-c", "--code", help="VBScript code to execute directly")

    args = parser.parse_args()

    try:
        if args.code:
            # Execute code from command line argument
            source = args.code
        elif args.file:
            # Read from file
            with open(args.file, "r") as f:
                source = f.read()
        else:
            # Read from stdin
            source = sys.stdin.read()

        if source.strip():
            run(source)
    except FileNotFoundError:
        print(f"Error: File not found: {args.file}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
