ECG Diagnostics Core v5.2.4 - pacote revisado com foco em CPU responsável e latência UNC

Conteúdo do pacote:
- ECG_Diagnostics_Core.ps1
- ECG_Diagnostics_Hub.bat
- ECG_FieldKit.ini
- README_RELEASE.txt
- SHA256SUMS.txt

Melhorias desta revisão:
- captura do processo dominante ficou mais robusta, com fallback adicional por contadores de performance
- processo dominante agora aparece de forma mais visível no HTML, nos eventos relevantes e no JSON
- adicionada medição de latência UNC por amostra para banco e NetDir
- benchmark passa a calcular média, p95 e pico das latências de banco e NetDir
- hipótese SHARE foi recalibrada para diferenciar "indisponível" de "acessível, porém lento"
- adicionada coleta opcional da CPU do processo ECGV6 via contador de performance
- adicionada coleta opcional de fila de disco total
- adicionada coleta opcional de bytes totais de rede
- INI ganhou novas chaves para ativar/desativar métricas opcionais
- versão do core, hub e perfil alinhada para v5.2.4

Implantação recomendada:
1. Copie os arquivos revisados para uma pasta de validação.
2. Ajuste o ECG_FieldKit.ini conforme o ambiente.
3. Execute o ECG_Diagnostics_Hub.bat.
4. Valide os modos Auto e Monitor em uma estação EXECUTANTE e uma VIEWER.
5. Para Fix e Rollback, execute o Hub elevado.

Saídas esperadas:
- Auto / Fix / Monitor:
  C:\ECG\FieldKit\out\<RunId>\ECG_Report.html
  C:\ECG\FieldKit\out\<RunId>\ECG_Report.json
  C:\ECG\FieldKit\out\<RunId>\ECG_Fatal_Error.log (somente em falha)

Observação honesta:
- esta revisão foi feita por análise de código e refatoração do gerador do laudo
- não houve execução runtime real em Windows dentro deste ambiente Linux
- a recomendação continua sendo validação controlada em Windows real antes de rollout amplo
