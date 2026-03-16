#!/usr/bin/env python3
"""
Generate a presentation with embedded video using officecli.

This script:
  1. Creates a short MP4 video (color gradient with animated bar) using imageio
  2. Creates a cover image (first frame) as PNG
  3. Builds a multi-slide PPTX with the video embedded

Requirements:
  pip install imageio imageio-ffmpeg numpy

Usage:
  python3 examples/gen-video-pptx.py
"""

import subprocess
import os
import sys
import tempfile
import shutil

def run(cmd):
    """Run officecli command and print it."""
    print(f"  $ officecli {cmd}")
    result = subprocess.run(f"officecli {cmd}", shell=True, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"    ERROR: {result.stderr.strip()}")
        sys.exit(1)
    if result.stdout.strip():
        print(f"    {result.stdout.strip()}")

def generate_video(video_path, cover_path):
    """Generate a 3-second 640x360 MP4 video and extract first frame as cover."""
    try:
        import imageio.v3 as iio
        import numpy as np
    except ImportError:
        print("ERROR: imageio not installed. Run: pip install imageio imageio-ffmpeg numpy")
        sys.exit(1)

    print("  Generating video frames...")
    W, H, FPS, DURATION = 640, 360, 30, 3
    total_frames = FPS * DURATION
    frames = []

    for i in range(total_frames):
        t = i / (total_frames - 1)
        frame = np.zeros((H, W, 3), dtype=np.uint8)

        # Gradient background: deep blue -> teal -> purple
        for y in range(H):
            yf = y / H
            r = int(20 + 60 * t + 40 * yf)
            g = int(30 + 80 * (1 - abs(t - 0.5) * 2) * (1 - yf))
            b = int(80 + 120 * (1 - t) + 50 * yf)
            frame[y, :, 0] = min(r, 255)
            frame[y, :, 1] = min(g, 255)
            frame[y, :, 2] = min(b, 255)

        # Moving circle
        cx = int(100 + t * (W - 200))
        cy = H // 2
        radius = 40
        yy, xx = np.ogrid[:H, :W]
        mask = (xx - cx) ** 2 + (yy - cy) ** 2 < radius ** 2
        frame[mask, 0] = 255
        frame[mask, 1] = 200
        frame[mask, 2] = 50

        # Text-like horizontal bars (simulating text lines)
        for row in range(3):
            bar_y = 60 + row * 50
            bar_w = int(200 + 100 * (1 - abs(t - 0.5) * 2))
            bar_x = 50
            frame[bar_y:bar_y + 12, bar_x:bar_x + bar_w, :] = [200, 200, 220]

        frames.append(frame)

    # Write video
    print(f"  Writing video: {video_path}")
    iio.imwrite(video_path, frames, fps=FPS)

    # Save first frame as cover
    print(f"  Writing cover: {cover_path}")
    iio.imwrite(cover_path, frames[0])


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    out_pptx = os.path.join(script_dir, f"{script_name}.pptx")

    # Create temp files for video and cover
    tmp_dir = tempfile.mkdtemp(prefix="officecli_video_")
    video_path = os.path.join(tmp_dir, "demo.mp4")
    cover_path = os.path.join(tmp_dir, "cover.png")

    try:
        # Step 1: Generate video and cover
        print("[1/4] Generating video and cover image...")
        generate_video(video_path, cover_path)
        video_size = os.path.getsize(video_path)
        print(f"  Video: {video_size / 1024:.1f} KB")

        # Step 2: Create presentation
        print(f"\n[2/4] Creating presentation: {out_pptx}")
        if os.path.exists(out_pptx):
            os.remove(out_pptx)
        run(f'create "{out_pptx}"')
        run(f'open "{out_pptx}"')

        # Slide 1 - Title slide with gradient background
        print("\n[3/4] Building slides...")
        print("  -- Slide 1: Title --")
        run(f'add "{out_pptx}" / --type slide --prop layout=title')
        run(f'set "{out_pptx}" /slide[1] --prop background=radial:1B2838-4472C4-bl')
        run(f'set "{out_pptx}" /slide[1]/placeholder[title] --prop text="Video Demo" --prop color=FFFFFF --prop size=44')
        run(f'set "{out_pptx}" /slide[1]/placeholder[subtitle] --prop text="Embedded video with officecli" --prop color=B4C7E7 --prop size=20')

        # Slide 2 - Video slide
        print("  -- Slide 2: Video --")
        run(f'add "{out_pptx}" / --type slide --prop title="Animated Video"')
        run(f'set "{out_pptx}" /slide[2] --prop background=0D1B2A')
        run(f'set "{out_pptx}" /slide[2]/shape[1] --prop color=FFFFFF')
        run(f'add "{out_pptx}" /slide[2] --type video '
            f'--prop path="{video_path}" '
            f'--prop poster="{cover_path}" '
            f'--prop x=2cm --prop y=4cm --prop width=22cm --prop height=12.5cm '
            f'--prop volume=80 --prop autoplay=true')

        # Slide 3 - Video info with chart
        print("  -- Slide 3: Video Stats --")
        run(f'add "{out_pptx}" / --type slide --prop title="Video Properties"')
        run(f'set "{out_pptx}" /slide[3] --prop background=1B2838')
        run(f'set "{out_pptx}" /slide[3]/shape[1] --prop color=FFFFFF')
        run(f'add "{out_pptx}" /slide[3] --type shape '
            f'--prop text="Resolution: 640x360\\nFPS: 30\\nDuration: 3s\\nFormat: MP4" '
            f'--prop font=Consolas --prop size=16 --prop color=B4C7E7 '
            f'--prop x=1cm --prop y=4cm --prop width=10cm --prop height=6cm '
            f'--prop fill=0D1B2A --prop line=4472C4 --prop linewidth=1pt')
        run(f'add "{out_pptx}" /slide[3] --type chart '
            f'--prop chartType=bar --prop title="Frame Colors" '
            f'--prop categories="Red,Green,Blue" '
            f'--prop "series1=Start:20,30,200" '
            f'--prop "series2=End:80,30,80" '
            f'--prop colors=E74C3C,27AE60 '
            f'--prop x=13cm --prop y=4cm --prop width=12cm --prop height=8cm')

        # Close resident and verify
        run(f'close "{out_pptx}"')

        print("\n[4/4] Verifying...")
        run(f'get "{out_pptx}" / --depth 1 --json')

        print(f"\nDone! Output: {out_pptx}")
        print(f"Open with: open \"{out_pptx}\"")

    finally:
        # Clean up temp files
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
