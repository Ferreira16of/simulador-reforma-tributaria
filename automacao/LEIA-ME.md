# Automação da simulação (scripts de apoio)

Scripts usados na rotina "atualizar tudo". Os caminhos (`/mnt/user-data/uploads/...`, `/tmp/w/...`, `/tmp/claude-0/s`) estão fixos nos scripts: ajustar para o ambiente da sessão (pasta das planilhas, pasta de trabalho) antes de rodar.

- `cache.py`, `build_bases.py`, `finalize.py`, `flow_extract.py`: leem as DFs das empresas e geram `bases.json` e `flow.json` (em `fonte/`).
- `atualizar_reducao_bases.py`: grava shareCP/shareLoc (redutores 44,2% e 70%) em bases.json.
- `atualizar_creditos_simulacao.py`: gera a 2.Creditos `_simulacao_vN` (zera o CAPEX das empresas com `semCapex` em bases.json).
- `atualizar_input_dre_rt.py`: gera a 1.Input `_automatizado_vN` e o Check_RT (precisa de `simulador/` ao lado com bases.json, motor.js, calc_cli.js).
- `atualizar_ppt.py` + `calc_cli.js`, `flow_cli.js`, `cenario_padrao.json`: atualiza o PPT.
- `chk_df.py`, `diffx.py`: conferências (motor x DF; diferença célula a célula entre versões de planilha).
- `dre_2027/`: DRE 2027 por empresa (`build.py`, `mot27.js`, `mk2.py`).
O motor (`motor.js`) e os fontes do HTML ficam em `fonte/`.
