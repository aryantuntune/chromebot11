"""Inspect the auto-collected CAPTCHA dataset and prepare/guide model fine-tuning.

The bot saves every CONFIRMED-correct CAPTCHA to ``output/captcha_dataset/`` as
``<label>_<hash>.png`` (the filename IS the verified label). That accumulates a
free, accurately-labelled training set. This tool:

  1. reports the dataset size + character coverage,
  2. sanity-checks the current solver against the dataset (confirms it loads),
  3. lays the data out for the ddddocr trainer and prints the exact steps to
     produce a fine-tuned model that the bot then auto-uses.

Run it:  .venv/Scripts/python.exe train_captcha.py

NOTE on actual neural training: ddddocr's *inference* is pip-installed here, but
TRAINING a new model uses the separate ``dddd_trainer`` toolkit (PyTorch), which
is a heavyweight, ideally-GPU step. So this script prepares the data and gives
you the recipe rather than training in-process. Drop the trainer's output
(``model.onnx`` + ``charsets.json``) into ``output/captcha_model/`` and the bot
picks it up automatically on the next run (see ``_get_ocr`` in run_bot.py).
"""
from __future__ import annotations

import collections
import pathlib
import re
import sys

DATASET = pathlib.Path("output/captcha_dataset")
MODEL_DIR = pathlib.Path("output/captcha_model")
LABEL_RE = re.compile(r"^([a-z0-9]+)_[0-9a-f]+\.png$", re.IGNORECASE)


def _samples():
    if not DATASET.is_dir():
        return []
    out = []
    for p in DATASET.glob("*.png"):
        m = LABEL_RE.match(p.name)
        if m:
            out.append((p, m.group(1).lower()))
    return out


def main() -> int:
    samples = _samples()
    print(f"Dataset: {DATASET}  ->  {len(samples)} labelled CAPTCHA(s)")
    if not samples:
        print("No samples yet. Run the bot (BOT_CAPTCHA_DATASET=1, the default) "
              "to collect some, then re-run this.")
        return 0

    # Character coverage / length distribution.
    chars = collections.Counter("".join(lbl for _, lbl in samples))
    lengths = collections.Counter(len(lbl) for _, lbl in samples)
    print("Lengths:", dict(sorted(lengths.items())))
    print("Char coverage ({} distinct):".format(len(chars)),
          "".join(sorted(chars)))

    # Sanity-check the current solver against a slice of the dataset.
    try:
        import ddddocr
        ocr = ddddocr.DdddOcr(show_ad=False)
        check = samples[:50]
        ok = 0
        for p, lbl in check:
            guess = re.sub(r"[^a-z0-9]", "", ocr.classification(p.read_bytes()).lower())
            ok += (guess == lbl)
        print(f"Current solver agreement on {len(check)} stored samples: "
              f"{ok}/{len(check)} ({ok/len(check)*100:.0f}%)")
        print("  (high by design -- these are samples the solver already got "
              "right; the value of the dataset is VOLUME + variety for training.)")
    except Exception as e:  # noqa: BLE001
        print("Solver check skipped:", e)

    # The dataset is already in dddd_trainer's expected '<label>_<id>.png' layout.
    MODEL_DIR.mkdir(parents=True, exist_ok=True)
    print()
    print("To fine-tune a model from this dataset:")
    print("  1. pip install the trainer toolkit:  git clone "
          "https://github.com/sml2h3/dddd_trainer")
    print(f"  2. point its dataset path at:  {DATASET.resolve()}")
    print("     (filenames are already 'label_id.png' -- the format it expects)")
    print("  3. train, then export to ONNX (the trainer's `export` step).")
    print(f"  4. copy the result into:  {MODEL_DIR.resolve()}")
    print("       model.onnx  +  charsets.json")
    print("  5. just run the bot again -- it auto-detects and uses that model.")
    print()
    print("Until then the bot uses the bundled model (already ~100% with the "
          "portal's unlimited CAPTCHA retries), and keeps growing this dataset.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
