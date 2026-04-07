# Setup | Sync automático das fichas para a matriz mestra

Este setup habilita sincronização automática de campos da ficha (Google Docs) para a planilha mestra.

## O que sincroniza
- `codigo_interno`
- `revisao`
- `vigencia`
- `status`
- `local_vigente`
- `retencao_minima`
- `ultima_atualizacao`
- `observacoes` (em caso de erro de leitura)

A chave de vínculo é `drive_doc_id`.
A coluna `vigencia` é criada automaticamente na planilha, se ainda não existir.

## Arquivo de script
- Código: `DocPackSync.gs`

## Passo a passo (Apps Script)
1. Abra o projeto Apps Script que você usa para automações internas.
2. Crie um novo arquivo `.gs` e cole o conteúdo de `DocPackSync.gs`.
3. Salve.
4. Execute uma vez a função `setupDocumentPackSync()` (sem parâmetros) para gravar a planilha padrão no script.

### Exemplo de configuração manual (opcional)
```javascript
function setupDocPackSync() {
  setDocumentPackSyncConfigFromUrl(
    'https://docs.google.com/spreadsheets/d/195Yps812fb5l2ftWD4lJ3zoAkExEZQyj3J7rZ185G4s/edit',
    '' // vazio = primeira aba
  );
}
```

5. Rode `runDocumentPackSyncNow()` para teste imediato.
6. Rode `createDocumentPackSyncTrigger()` para criar trigger diário (aprox. 02:00).

## Funções úteis
- `runDocumentPackSyncNow()` → sincronização manual sob demanda
- `createDocumentPackSyncTrigger()` → cria trigger diário (aprox. 02:00)
- `createDocumentPackSyncTriggerHourly()` → cria trigger horário (se precisar maior frequência)
- `clearDocumentPackSyncTriggers()` → remove trigger(s)

## Botão / menu para rodar a sync
- Ao abrir a planilha, o script cria um menu: **Document Pack Sync**.
- Use **Document Pack Sync > Run sync now** para rodar manualmente com um clique.
- Também há opções no menu para setup e triggers.

### Se quiser um botão visual na planilha
1. Na planilha Google Sheets, vá em **Inserir > Desenho** ou **Inserir > Imagem**.
2. Crie um botão visual como “Run Doc Sync”.
3. Clique nos três pontos do objeto e use **Atribuir script**.
4. Informe o nome: `runDocumentPackSyncFromMenu`

## Observações importantes
- O usuário dono do trigger precisa ter acesso de leitura/edição nas fichas (Docs) e edição na planilha.
- Como Google Docs não tem trigger nativo de edição equivalente ao Sheets, o modelo recomendado é **time-driven** (polling).
- O padrão deste projeto está em modo diário; use modo horário apenas se houver necessidade operacional.
