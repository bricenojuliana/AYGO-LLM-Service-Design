"""Renders the icons (SVG -> PNG) and the Graphviz diagrams in diagramas/.

Needs cairosvg (in requirements.txt) and the `dot` command from Graphviz.
"""
import subprocess
from pathlib import Path

import cairosvg

DIAGRAMS = Path(__file__).resolve().parent.parent / "diagramas"


def main():
    for svg in sorted((DIAGRAMS / "iconos").glob("*.svg")):
        cairosvg.svg2png(url=str(svg), write_to=str(svg.with_suffix(".png")), output_width=96, output_height=96)
    for dot in sorted(DIAGRAMS.glob("*.dot")):
        out = dot.with_suffix(".png")
        subprocess.run(["dot", "-Tpng", "-Gdpi=170", dot.name, "-o", out.name], cwd=DIAGRAMS, check=True)
        print(f"Saved diagramas/{out.name}")


if __name__ == "__main__":
    main()
