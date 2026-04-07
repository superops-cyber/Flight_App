# SOP | Controle de Formulários do Document Pack (RBAC 135)

## 1. Objetivo
Padronizar numeração interna, finalidade, acesso, armazenamento vigente e arquivamento dos formulários de suporte operacional e técnico do pacote documental RBAC 135.

## 2. Escopo
Aplica-se aos 40 formulários listados na matriz mestra em `135/document_pack_master_index.csv`.

## 3. Convenção de numeração interna
- Formato: **FRM-EXT-XXX**
- `XXX`: número sequencial de 3 dígitos baseado no prefixo do arquivo local (ex.: `01_...` = `FRM-EXT-001`)
- Exemplo: `135/document_pack/21_lista_de_passageiros.md` → `FRM-EXT-021`

## 4. Campos obrigatórios da matriz
A matriz mestra deve manter, no mínimo, os campos abaixo para governança documental:
- `codigo_interno`
- `uso_do_formulario`
- `acesso_requerido`
- `local_vigente`
- `local_arquivo_morto`
- `retencao_minima`

## 5. Níveis de acesso (acesso_requerido)
- **PILOTO + DISPATCH + BASE**: formulários de uso direto em missão (ex.: passageiros, W&B, manifesto)
- **DISPATCH + PILOTO (consulta)**: documentos de liberação e decisão operacional
- **DISPATCH + BASE**: documentos de rotina operacional não embarcados
- **DISPATCH + MANUTENÇÃO**: controle de aeronave e prontidão técnica
- **MANUTENÇÃO + GESTÃO**: documentos de manutenção/OM145
- **GESTÃO + QUALIDADE**: governança, regulatório, programas e controle sistêmico

## 6. Regras de armazenamento
### 6.1 Vigente
- Operações com acesso de piloto: `Drive/OPS/Formularios_Vigentes` + cópia a bordo quando aplicável
- Aeronave: `Drive/OPS/Aeronaves/Formularios_Vigentes`
- Manutenção: `Drive/MNT/Controle_Tecnico/Formularios_Vigentes`
- Governança/Regulatório/Programas: `Drive/SGQ/Documentos_Controlados/Formularios_Vigentes`

### 6.2 Arquivo morto
- Operações: `Drive/OPS/Arquivo_Morto` (com via física na base quando aplicável)
- Aeronave/Governança/Regulatório/Programas: `Drive/SGQ/Arquivo_Morto` + backup NAS
- Manutenção: `Drive/MNT/Arquivo_Morto` + backup NAS_MNT + pasta física OM145

## 7. Retenção mínima
- Padrão: **36 meses**
- Registros operacionais críticos (jornada, fadiga, treinamento, manifesto, W&B, passageiros, despacho): **60 meses**
- Contratos: **60 meses após término contratual**

## 8. Responsabilidades
- **DOPS/Dispatch**: garantir versão vigente e disponibilidade operacional
- **Qualidade/SGQ**: controlar revisão, distribuição e rastreabilidade
- **Manutenção**: custodiar registros técnicos e evidências de aeronavegabilidade
- **Base/PIC**: manter documentação de bordo aplicável e devolver para arquivamento

## 9. Fluxo resumido
1. Emitir/atualizar formulário vigente no local oficial.
2. Aplicar número interno (`FRM-EXT-XXX`) e revisão.
3. Controlar acesso conforme coluna `acesso_requerido`.
4. Encerrar uso operacional e mover para `local_arquivo_morto`.
5. Manter retenção mínima conforme coluna `retencao_minima`.
