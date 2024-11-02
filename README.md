# pptx-png-text-label-extractor

Simple way to get those annoying icon PNGs and associated text into your grateful hands from a PPTX file

## Get your Python environment ready

```
johnmaeda~> python -m venv venv
johnmaeda~> source venv/bin/activate
(venv) johnmaeda~> python -m pip install python-pptx Pillow
```

## You're ready to roll

This will take all the pages in the sample.pptx and output all the icons into an `out` directory. Take note that they need to be PNGs.

```
python extract.py --help
python extract.py sample.pptx
```

And to grab all the icons that have poured into your `out` directory so they're available in a single PPT file, just:

```
python restore.py
```
