# Drift notes detectados no pacote local v5.2.1

## 1. SHA256SUMS.txt

O arquivo de checksums carregado junto do pacote não corresponde aos hashes atuais dos arquivos locais usados nesta sessão.

Por isso, este bootstrap publica `SHA256SUMS_CORRETOS.txt` como referência canônica para a extração do novo repositório.

## 2. Version strings internas

Há sinais de versionamento misto em partes textuais do pacote (comentários/cabeçalhos antigos convivendo com banner e release atualizados).

Recomendação:
- alinhar todos os banners textuais para `v5.2.1`
- depois gerar o checksum final do pacote definitivo
- somente então publicar a primeira release do novo repositório
