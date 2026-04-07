# Política de Identificação e Etiquetagem de Peças

Versão: 0.1
Data: 2026-03-16

## 1. Objetivo

Definir o sistema mínimo de identificação, etiquetagem e rastreabilidade de peças, componentes e conjuntos utilizados, removidos ou armazenados durante atividades de manutenção, garantindo evidência documental, segregação de peças não conformes e prevenção de instalação inadvertida.

## 2. Alcance

Aplica‑se a todas as peças e componentes geridos pelo operador e pelas OM RBAC‑145 contratadas, incluindo peças recebidas de fornecedores, peças em stock, peças removidas para manutenção e peças em investigação ou descarte.

## 3. Código de cores e significados

- Verde — Serviceable/Ready: peça com documentação completa (certificado, PN/SN): pronta para instalação.
- Amarelo — Quarentena: peça pendente de verificação documental, inspeção ou certificação; instalar apenas com autorização escrita do DOPS ou do certificador.
- Vermelho — Não‑serviceable: peça condenada; não instalar; destino: overhaul/disposição.
- Azul — Retida para investigação: peça retida por investigação, garantia ou análise; não instalar até resolução.

## 4. Informação mínima da etiqueta

- Identificador único da etiqueta / código (ID / QR)
- PN (Part Number)
- SN (Serial Number), quando aplicável
- Quantidade e unidade
- Aeronave (matrícula) ou fornecedor de origem
- Referência do workcard / serviço ou número de recebimento
- Data / hora da ação
- Motivo da etiqueta (ex.: removida, quarentena documental, condenada)
- Estado sugerido (reinstalar / overhaul / descartar)
- Responsável / assinatura

## 5. Procedimento operacional

1. No ato da remoção, recepção ou triagem a peça deverá receber a etiqueta com o código de cor correspondente.
2. Peças pequenas deverão ser acondicionadas em saco ou envelope resistente com a etiqueta afixada externamente.
3. Peças sem identificação legível serão imediatamente colocadas em quarentena (Amarelo) até investigação.
4. Peças condenadas serão segregadas e etiquetadas em Vermelho e tratadas conforme procedimento de descarte do operador.
5. Todas as alterações de estado da peça (por ex.: quarentena → serviceable) deverão estar documentadas no workcard e aprovadas por responsável autorizado.

## 6. Registos e integração com sistema de manutenção

- Fotografar a peça etiquetada e anexar a imagem ao workcard.
- Registrar o ID da etiqueta e os campos mínimos no sistema eletrônico de manutenção (ou em formulário controlado) para assegurar rastreabilidade.

## 7. Responsabilidades

- OM RBAC‑145 / Oficina: aplicar etiquetas, manter documentação de origem e enviar evidências ao operador.
- DOPS: autorizar instalações excecionais de peças em quarentena e validar retornos ao serviço.
- GR: assegurar fornecimento de etiquetas, conformidade contratual e supervisão de processos.

## 8. Auditoria e formação

- Incluir o processo de etiquetagem nas auditorias periódicas da OM.
- Realizar formação anual para técnicos sobre o uso correto das etiquetas e reconciliação de stock.

## 9. Modelos (exemplo de tag)

Tag ID: ____________  Cor: ( ) Verde  ( ) Amarelo  ( ) Vermelho  ( ) Azul
PN: ____________________  SN: ____________________  Qtd: _______
Origem / Aeronave / Fornecedor: ___________________________
Workcard / Recebimento: ____________________
Data / Hora: ____________  Responsável: ____________________
Motivo / Observações:
_______________________________________________________________
Disposição prevista: ( ) Reinstalar  ( ) Overhaul  ( ) Descartar

---

Arquivo adicional: fluxograma de fluxo de peças e modelo de etiqueta eletrónica disponíveis no índice mestre e no arquivo `maintenance/`.
