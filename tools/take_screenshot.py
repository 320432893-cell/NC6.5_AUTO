import argparse
import os

from core.jab_operator import take_desktop_screenshot


def main():
    parser = argparse.ArgumentParser(description="Take a desktop screenshot.")
    parser.add_argument("--output", default="nc_screen.png")
    args = parser.parse_args()
    if os.name != "nt":
        raise SystemExit("This tool must run with Windows Python.")

    image = take_desktop_screenshot()
    image.save(args.output)
    print({"output": args.output, "size": image.size})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
