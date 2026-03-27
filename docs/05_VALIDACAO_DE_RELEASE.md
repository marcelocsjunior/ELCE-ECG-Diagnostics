# 05_VALIDACAO_DE_RELEASE

## Checklist oficial

### Contrato BAT + PS1
- [ ] BAT aponta para o PS1 correto
- [ ] Não há mistura de linhas diferentes

### Execução
- [ ] O menu abre normalmente
- [ ] A opção de diagnóstico executa sem quebra
- [ ] O `RunId` é criado corretamente

### Artefatos
- [ ] `execution.log` gerado
- [ ] `context.json` gerado
- [ ] `timeline.json` gerado
- [ ] `benchmark.json` gerado
- [ ] `analysis.json` gerado
- [ ] `ELCE_ECG_Diagnostics_Report.html` gerado
- [ ] summaries gerados sem afetar o fechamento da rodada

### Publicação
- [ ] `Latest` atualizado corretamente
- [ ] `latest_run.txt` atualizado corretamente

### Qualidade
- [ ] HTML abre sem erro
- [ ] Conteúdo principal do laudo está legível
- [ ] Acentuação verificada
- [ ] Gráfico validado
- [ ] Regressão operacional não observada

## Regra final

Sem checklist completo, não há promoção de baseline.
