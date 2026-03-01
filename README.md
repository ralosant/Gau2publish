
# gau2publish

Convert Gaussian (`.log`) outputs into:
- `.geom`: ordered coordinates with metadata
- `.png`: 3D render (if `xyzrender` is available; can be disabled with `--no-xyzrender`)
- `.cdxml`: ChemDraw document (if `pycdxml` is available)
- `DOCX`: a document with title, image and the `.geom` block
- `ALL.cdxml`: a slide-like mosaic with thumbnails and properties

## Installation

### pip
```bash
pip install .
```

### conda (conda-forge)
Once the recipe is merged:
```bash
conda install -c conda-forge gau2publish
```

## Usage
```bash
gau2publish *.log --docx ESI_xyz.docx --method wB97XD   --image-width 6 --title-size 14 --slide-columns 6   --xyzrender-args "-I -S 1200"
```

## Notes
- `xyzrender` is optional. Use `--no-xyzrender` to skip PNG generation.
- If `pycdxml` is not installed, `.cdxml` and `ALL.cdxml` generation will be skipped.
- `Openbabel` is used to convert `.log`→`.sdf` and to extract properties.
