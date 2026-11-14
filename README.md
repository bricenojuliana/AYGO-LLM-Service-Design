# Designing an LLM Service That Can Grow

AYGO lab. The goal is an infrastructure proposal for a Spanish-language chat assistant serving users in Latin America: demand, model, memory, capacity, architecture, location, power, cooling, cost and failure handling.

The deliverables themselves are written in Spanish. This repo only holds them plus the scripts that build them.

## What's here

| Path | Content |
|------|---------|
| `modelo/modelo_calculo.xlsx` | Calculation model (inputs, decision log, and later demand, memory, inventory, power, costs, scenarios) |
| `propuesta/` | Technical proposal (Word), added in later steps; `notas_secciones.md` holds working notes per section |
| `diagramas/` | Graphviz sources (`.dot`), rendered diagrams and hand-drawn line icons |
| `scripts/` | Python scripts that generate the Excel, the diagrams and the Word file |

## Rebuilding the files

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python scripts/build_model.py
.venv/bin/python scripts/build_diagrams.py   # needs Graphviz (`dot`)
```

In the Excel file, blue values are inputs typed by hand. Everything else should come from formulas, so changing an input updates the rest.

## Status

- [x] Step 0: project setup, inputs and decision log
- [x] Step 1: demand (requests, tokens, concurrency, sensitivity, 2x stress)
- [x] Step 2: model and memory (Llama 3.1 70B BF16 vs FP8 on H100)
- [x] Step 3: performance (synthetic 2 req/s per replica + order-of-magnitude checks)
- [x] Step 4: replicas, nodes and headroom (9 nodes, 72 GPUs, 18 replicas)
- [x] Step 5: service architecture (logical and deployment diagrams, overload sheet)
- [x] Step 6: networking and storage (traffic map, storage budget)
- [x] Step 7: location (colocation in São Paulo vs Querétaro and Santiago; exotic options ruled out)
- [x] Step 8: power, energy and cooling (normal, maximum and degraded balances)
- [x] Step 9: budget (CAPEX, OPEX, TCO and 8 cost scenarios)
- [ ] Step 10: resilience and recovery
