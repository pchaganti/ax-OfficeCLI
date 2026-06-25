#!/usr/bin/env python3
"""
Animation Showcase — generates animations.pptx demonstrating the pptx animation
and slide-transition vocabulary: entrance / exit / emphasis effects (the
`animation=<effect>-<class>-<ms>[-with|-after]` shape prop), the slide-level
`transition=` gallery, and timing/trigger variants (click / after / with,
slow vs fast, auto-advance).

SDK twin of animations.sh (officecli CLI). Both produce an equivalent
animations.pptx. This one drives the **officecli Python SDK**
(`pip install officecli-sdk`): one resident is started and every slide, shape,
and set is shipped over the named pipe in `doc.batch(...)` round-trips. Each
item is the same `{"command","parent"/"path","type","props"}` dict you'd put in
an `officecli batch` list.

Usage:
  pip install officecli-sdk          # plus the `officecli` binary on PATH
  python3 animations.py
"""

import os
import sys

# --- locate the SDK: prefer an installed `officecli-sdk`, else the in-repo copy
try:
    import officecli  # pip install officecli-sdk
except ImportError:
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "..", "..", "sdk", "python"))
    import officecli

FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "animations.pptx")


def add_slide(**props):
    """One `add slide` item in batch-shape."""
    return {"command": "add", "parent": "/", "type": "slide", "props": props}


def add_shape(slide, **props):
    """One `add shape` item in batch-shape, targeting /slide[N]."""
    return {"command": "add", "parent": f"/slide[{slide}]", "type": "shape", "props": props}


def setp(path, **props):
    """One `set` item in batch-shape."""
    return {"command": "set", "path": path, "props": props}


print(f"Building {FILE} ...")

with officecli.create(FILE, "--force") as doc:

    # =====================================================================
    # SLIDE 1 — Title
    # =====================================================================
    print("  -> Slide 1: Title")
    doc.batch([
        add_slide(layout="title"),
        setp("/slide[1]", background="radial:0D1B2A-1B4F72-bl"),
        setp("/slide[1]/placeholder[centertitle]",
             text="Animation Showcase", color="FFFFFF", size="48"),
        setp("/slide[1]/placeholder[subtitle]",
             text="Every animation effect in officecli", color="85C1E9", size="22"),
        setp("/slide[1]", transition="fade"),
    ])

    # =====================================================================
    # SLIDE 2 — Entrance Animations
    # =====================================================================
    print("  -> Slide 2: Entrance Animations")
    # (fill, preset, x, y, width, height, size, animation) per demo shape;
    # shapes land at /slide[2]/shape[2..13] (shape[1] is the title placeholder).
    entrances = [
        ("appear",       "2E86C1", "1cm",  "4cm",   "14", "appear-entrance-500"),
        ("fade",         "27AE60", "7cm",  "4cm",   "14", "fade-entrance-800"),
        ("fly",          "E74C3C", "13cm", "4cm",   "14", "fly-entrance-600"),
        ("zoom",         "8E44AD", "19cm", "4cm",   "14", "zoom-entrance-700"),
        ("wipe",         "F39C12", "1cm",  "7.5cm", "14", "wipe-entrance-600"),
        ("bounce",       "1ABC9C", "7cm",  "7.5cm", "14", "bounce-entrance-800"),
        ("float",        "E67E22", "13cm", "7.5cm", "14", "float-entrance-700"),
        ("split",        "2980B9", "19cm", "7.5cm", "14", "split-entrance-600"),
        ("wheel",        "C0392B", "1cm",  "11cm",  "14", "wheel-entrance-800"),
        ("swivel",       "16A085", "7cm",  "11cm",  "14", "swivel-entrance-700"),
        ("checkerboard", "D35400", "13cm", "11cm",  "12", "checkerboard-entrance-600"),
        ("blinds",       "7D3C98", "19cm", "11cm",  "14", "blinds-entrance-600"),
    ]
    items = [
        add_slide(title="Entrance Effects"),
        setp("/slide[2]", background="1B2838"),
        setp("/slide[2]/shape[1]", color="FFFFFF", size="28"),
    ]
    for idx, (text, fill, x, y, size, anim) in enumerate(entrances):
        items.append(add_shape(2, text=text, font="Consolas", size=size, color="FFFFFF",
                               fill=fill, preset="roundRect",
                               x=x, y=y, width="5cm", height="2cm"))
        items.append(setp(f"/slide[2]/shape[{idx + 2}]", animation=anim))
    items.append(setp("/slide[2]", transition="wipe"))
    doc.batch(items)

    # =====================================================================
    # SLIDE 3 — Exit Animations
    # =====================================================================
    print("  -> Slide 3: Exit Animations")
    exits = [
        ("fade out",     "E74C3C", "1cm",  "4cm", "fade-exit-800"),
        ("fly out",      "2E86C1", "9cm",  "4cm", "fly-exit-600"),
        ("zoom out",     "27AE60", "17cm", "4cm", "zoom-exit-700"),
        ("dissolve out", "8E44AD", "1cm",  "8cm", "dissolve-exit-600"),
        ("wipe out",     "F39C12", "9cm",  "8cm", "wipe-exit-600"),
        ("flash out",    "1ABC9C", "17cm", "8cm", "flash-exit-500"),
    ]
    items = [
        add_slide(title="Exit Effects"),
        setp("/slide[3]", background="1B2838"),
        setp("/slide[3]/shape[1]", color="FFFFFF", size="28"),
    ]
    for idx, (text, fill, x, y, anim) in enumerate(exits):
        items.append(add_shape(3, text=text, font="Consolas", size="14", color="FFFFFF",
                               fill=fill, preset="roundRect",
                               x=x, y=y, width="7cm", height="2.5cm"))
        items.append(setp(f"/slide[3]/shape[{idx + 2}]", animation=anim))
    items.append(setp("/slide[3]", transition="push"))
    doc.batch(items)

    # =====================================================================
    # SLIDE 4 — Emphasis Animations
    # =====================================================================
    print("  -> Slide 4: Emphasis Animations")
    emphases = [
        ("spin", "E74C3C", "2cm",  "spin-emphasis-1000"),
        ("grow", "2E86C1", "8cm",  "grow-emphasis-800"),
        ("wave", "27AE60", "14cm", "wave-emphasis-700"),
        ("bold", "8E44AD", "20cm", "bold-emphasis-500"),
    ]
    items = [
        add_slide(title="Emphasis Effects"),
        setp("/slide[4]", background="1B2838"),
        setp("/slide[4]/shape[1]", color="FFFFFF", size="28"),
    ]
    for idx, (text, fill, x, anim) in enumerate(emphases):
        items.append(add_shape(4, text=text, font="Consolas", size="16", color="FFFFFF",
                               fill=fill, preset="ellipse",
                               x=x, y="4.5cm", width="4.5cm", height="4.5cm"))
        items.append(setp(f"/slide[4]/shape[{idx + 2}]", animation=anim))
    items.append(setp("/slide[4]", transition="zoom"))
    doc.batch(items)

    # =====================================================================
    # SLIDE 5 — Slide Transitions Gallery
    # =====================================================================
    print("  -> Slide 5: Transitions Gallery")
    transitions = ["fade", "wipe", "push", "split", "zoom", "wheel", "cover",
                   "reveal", "dissolve", "random", "blinds", "checker", "strips"]
    items = [
        add_slide(title="Slide Transitions"),
        setp("/slide[5]", background="0D1B2A"),
        setp("/slide[5]/shape[1]", color="FFFFFF", size="28"),
    ]
    # 4-column grid: x = 1 + col*6 cm, y steps +2.5cm every 4 shapes (matches
    # the shell's X=1/COL*6 and Y=4 then +2.5 loop).
    col = 0
    y = 4.0
    for tr in transitions:
        px = 1 + col * 6
        items.append(add_shape(5, text=tr, font="Consolas", size="12", color="FFFFFF",
                               fill="2C3E50", preset="roundRect",
                               line="5DADE2", linewidth="0.5pt",
                               x=f"{px}cm", y=f"{y}cm", width="5cm", height="1.8cm"))
        col += 1
        if col >= 4:
            col = 0
            y += 2.5
    items.append(setp("/slide[5]", transition="split"))
    doc.batch(items)

    # =====================================================================
    # SLIDE 6 — Timing & Triggers
    # =====================================================================
    print("  -> Slide 6: Timing & Triggers")
    triggers = [
        ("Click to animate\n(default trigger)", "2E86C1", "1cm",  "4cm", "7cm",  "3cm", "fade-entrance-500"),
        ("After previous\n(auto-follows)",      "27AE60", "9cm",  "4cm", "7cm",  "3cm", "fly-entrance-500-after"),
        ("With previous\n(simultaneous)",       "E74C3C", "17cm", "4cm", "7cm",  "3cm", "zoom-entrance-500-with"),
        ("Slow (2000ms)",                       "8E44AD", "1cm",  "9cm", "11cm", "3cm", "wipe-entrance-2000"),
        ("Fast (200ms)",                        "F39C12", "13cm", "9cm", "11cm", "3cm", "wipe-entrance-200"),
    ]
    items = [
        add_slide(title="Timing & Triggers"),
        setp("/slide[6]", background="1B2838"),
        setp("/slide[6]/shape[1]", color="FFFFFF", size="28"),
    ]
    for idx, (text, fill, x, y, w, h, anim) in enumerate(triggers):
        items.append(add_shape(6, text=text, font="Consolas", size="13", color="FFFFFF",
                               fill=fill, preset="roundRect",
                               x=x, y=y, width=w, height=h))
        items.append(setp(f"/slide[6]/shape[{idx + 2}]", animation=anim))
    items.append(setp("/slide[6]", transition="reveal"))
    items.append(setp("/slide[6]", advanceTime="5000"))
    doc.batch(items)

print(f"Generated: {FILE}")
