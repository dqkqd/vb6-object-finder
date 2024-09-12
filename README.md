# VisualBasic6 find implicit cast object

## How to run

```python
python -m venv .venv
source .venv/bin/activate
pip install .
```

## Generate parser

```bash
antlr4 -Dlanguage=Python3 VisualBasic6.g4
```

## Sample

```python
python check.py
```
