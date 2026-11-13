# Designing an LLM Service That Can Grow

AYGO lab. The goal is an infrastructure proposal for a Spanish-language chat assistant serving users in Latin America: demand, model, memory, capacity, architecture, location, power, cooling, cost and failure handling.

The deliverables themselves are written in Spanish. This repo only holds them plus the scripts that build them.

## What's here

| Path | Content |
|------|---------|
| `modelo/modelo_calculo.xlsx` | Calculation model (inputs, decision log, and later demand, memory, inventory, power, costs, scenarios) |
| `propuesta/` | Technical proposal (Word), added in later steps; `notas_secciones.md` holds working notes per section |
| `scripts/` | Python scripts that generate the Excel and Word files |

## Rebuilding the files

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python scripts/build_model.py
```

In the Excel file, blue values are inputs typed by hand. Everything else should come from formulas, so changing an input updates the rest.

## Status

- [x] Step 0: project setup, inputs and decision log
- [x] Step 1: demand (requests, tokens, concurrency, sensitivity, 2x stress)
- [x] Step 2: model and memory (Llama 3.1 70B BF16 vs FP8 on H100)
- [ ] Step 3: performance evidence
