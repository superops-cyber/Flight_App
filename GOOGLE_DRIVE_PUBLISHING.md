# Publicar templates no Google Drive (Docs + índice)

Este projeto já tem os templates em:
- `135/document_pack/`

E agora tem um publicador automático:
- `publish_document_pack_to_drive.py`

## O que o script faz

1. Lê todos os `.md` em `135/document_pack/` (exceto `INDEX.md`).
2. Cria uma pasta raiz no Google Drive (ex.: `Asas - Document Pack`).
3. Cria subpasta `Google Docs`.
4. Converte cada template para **Google Docs** (um arquivo por documento).
5. Gera CSV mestre local com links clicáveis:
   - `135/document_pack_master_index.csv`
6. Sobe o CSV para o Drive convertido em **Google Sheets**.

## Pré-requisitos

- Python 3.10+
- Conta Google com acesso ao Drive
- OAuth Client JSON do Google Cloud (tipo Desktop App)

Instale dependências:

`pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib`

## Passo a passo (primeira vez)

1. No Google Cloud Console:
   - Ative Google Drive API
   - Crie credenciais OAuth Client ID (Desktop App)
   - Baixe `credentials.json`
2. Salve `credentials.json` na raiz do projeto.
3. Rode o script:

`python3 publish_document_pack_to_drive.py --credentials credentials.json --root-folder-name "Asas - Document Pack"`

### Publicar dentro de uma pasta compartilhada específica

Se você já tem uma pasta do projeto no Shared Drive, use o URL dela:

`python3 publish_document_pack_to_drive.py --credentials credentials.json --parent-folder-url "https://drive.google.com/drive/folders/SEU_FOLDER_ID" --root-folder-name "Asas - Document Pack"`

Se quiser publicar diretamente na pasta compartilhada (sem criar subpasta raiz nova):

`python3 publish_document_pack_to_drive.py --credentials credentials.json --parent-folder-url "https://drive.google.com/drive/folders/SEU_FOLDER_ID" --use-parent-as-root`

4. Na primeira execução, abrirá o navegador para login/autorização.
5. O script salva `token.json` para as próximas execuções.

## Dry run (sem tocar no Drive)

`python3 publish_document_pack_to_drive.py --dry-run`

Gera só o CSV local, sem upload.

## Resultado esperado

- Pasta no Drive com os documentos Google Docs
- Planilha Google com índice mestre e links
- CSV local atualizado com colunas:
  - `drive_doc_id`
  - `drive_doc_url`
  - `hyperlink_formula`

## Observações

- Eu não consigo autenticar na sua conta Google daqui; você precisa executar o script localmente para concluir o upload.
- Se quiser, posso adaptar o script para criar também subpastas por categoria (Governança, Operações, Aeronave, Manutenção, etc.).
