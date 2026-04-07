from pathlib import Path
import re

ROOT = Path('135')
CHECKLIST = ROOT / 'required_support_documents_checklist.md'
OUT_DIR = ROOT / 'document_pack'


def slugify(text: str) -> str:
    text = text.lower().strip()
    text = re.sub(r'\(.*?\)', '', text)
    text = text.replace('/', ' ')
    text = re.sub(r'[^a-z0-9\s_-]', '', text)
    text = re.sub(r'\s+', '_', text)
    text = re.sub(r'_+', '_', text).strip('_')
    return text[:80] if len(text) > 80 else text


def parse_rows(markdown: str):
    rows = []
    lines = markdown.splitlines()
    start = None
    for i, line in enumerate(lines):
        if line.strip().startswith('| Categoria | Documento extra |'):
            start = i + 2
            break
    if start is None:
        return rows

    for line in lines[start:]:
        if not line.strip().startswith('|'):
            break
        parts = [p.strip() for p in line.strip().strip('|').split('|')]
        if len(parts) != 7:
            continue
        categoria, documento, explicacao, exemplo, fonte, responsavel, prioridade = parts
        rows.append({
            'categoria': categoria,
            'documento': documento,
            'explicacao': explicacao,
            'exemplo': exemplo,
            'fonte': fonte,
            'responsavel': responsavel,
            'prioridade': prioridade,
        })
    return rows


def make_table(headers, rows) -> str:
    out = ['| ' + ' | '.join(headers) + ' |', '|' + '|'.join(['---'] * len(headers)) + '|']
    for row in rows:
        out.append('| ' + ' | '.join(row) + ' |')
    return '\n'.join(out)


def section(title: str, body: str) -> str:
    return f'## {title}\n\n{body.strip()}\n'


def common_control_table(row: dict) -> str:
    return make_table(
        ['Campo', 'Preenchimento'],
        [
            ['Documento', row['documento']],
            ['Categoria', row['categoria']],
            ['Prioridade', row['prioridade']],
            ['Responsável primário', row['responsavel']],
            ['Fonte (manuais)', row['fonte']],
            ['Código interno', '____________________'],
            ['Revisão', '____'],
            ['Vigência', '____/____/______'],
            ['Status', '☐ Em elaboração  ☐ Em revisão  ☐ Vigente  ☐ Obsoleto'],
            ['Substituto / backup', '____________________'],
            ['Local oficial de armazenamento', '____________________'],
            ['Prazo de retenção', '____________________'],
        ],
    )


def purpose_table(row: dict) -> str:
    return make_table(
        ['Item', 'Conteúdo'],
        [
            ['Finalidade', row['explicacao']],
            ['Exemplo de preenchimento', row['exemplo']],
            ['Regra de uso', 'Preencher os campos aplicáveis, obter aprovações e arquivar no local oficial.'],
            ['Critério de fechamento', 'Documento somente concluído após validação, assinatura e registro no índice mestre.'],
        ],
    )


def approval_table() -> str:
    return make_table(
        ['Papel', 'Nome', 'Assinatura / aceite', 'Data'],
        [
            ['Elaborado por', '____________________', '____________________', '____/____/______'],
            ['Revisado por', '____________________', '____________________', '____/____/______'],
            ['Aprovado por', '____________________', '____________________', '____/____/______'],
        ],
    )


def revision_table() -> str:
    return make_table(
        ['Revisão', 'Data', 'Descrição da alteração', 'Responsável'],
        [
            ['00', '____/____/______', 'Emissão inicial', '____________________'],
            ['__', '____/____/______', '____________________', '____________________'],
            ['__', '____/____/______', '____________________', '____________________'],
        ],
    )


def validation_list() -> str:
    return '\n'.join([
        '- ☐ Revisão vigente confirmada.',
        '- ☐ Campos obrigatórios preenchidos.',
        '- ☐ Evidências anexadas quando aplicável.',
        '- ☐ Aprovações e assinaturas obtidas.',
        '- ☐ Arquivamento concluído no local oficial.',
    ])


def generic_registry_form(headers, rows, checks=None) -> str:
    parts = [make_table(headers, rows)]
    if checks:
        parts.append('### Verificações obrigatórias\n\n' + '\n'.join(f'- ☐ {item}' for item in checks))
    return '\n\n'.join(parts)


def generic_summary_fields(pairs) -> str:
    return make_table(['Campo', 'Valor'], pairs)


def form_body(row: dict) -> str:
    n = row['documento'].lower()
    category = row['categoria'].lower()

    if 'cadastro de funções requeridas' in n:
        return generic_registry_form(
            ['Função regulatória', 'Titular', 'Suplente', 'Ato de designação', 'Início vigência', 'Fim vigência', 'Contato'],
            [
                ['Gestor Responsável', '', '', '', '', '', ''],
                ['Diretor de Operações', '', '', '', '', '', ''],
                ['Responsável SGSO', '', '', '', '', '', ''],
                ['Coordenador de Manutenção', '', '', '', '', '', ''],
                ['Outro cargo requerido', '', '', '', '', '', ''],
            ],
            ['Todos os cargos regulatórios estão preenchidos.', 'Atos de designação anexados ao processo.'],
        )

    if 'delegações e substituições' in n:
        return generic_registry_form(
            ['Data início', 'Data fim', 'Função titular', 'Substituto', 'Escopo delegado', 'Instrumento', 'Observações'],
            [['', '', '', '', '', '', ''] for _ in range(6)],
        )

    if 'autorizados para controle operacional' in n:
        return generic_registry_form(
            ['Nome', 'Função', 'Pode liberar voo', 'Pode suspender voo', 'Base', 'Contato', 'Vigência'],
            [['', '', '☐', '☐', '', '', ''] for _ in range(8)],
        )

    if 'índice mestre documental' in n:
        return generic_registry_form(
            ['Código', 'Documento', 'Categoria', 'Revisão', 'Status', 'Local oficial', 'Dono'],
            [['', '', '', '', '', '', ''] for _ in range(10)],
        )

    if 'controle de revisões e distribuição' in n:
        return generic_registry_form(
            ['Documento', 'Revisão', 'Destinatário', 'Meio', 'Ciência obtida', 'Data', 'Observações'],
            [['', '', '', '', '☐', '', ''] for _ in range(10)],
        )

    if 'lpe ou controle eletrônico equivalente' in n:
        return generic_registry_form(
            ['Manual', 'Revisão', 'Páginas / seções vigentes', 'Data da atualização', 'Responsável', 'Local'],
            [['', '', '', '', '', ''] for _ in range(6)],
        )

    if 'pacote eo/coa/sei vigente' in n:
        return generic_registry_form(
            ['Documento regulatório', 'Identificação', 'Revisão / data', 'Status', 'Local do arquivo', 'Observações'],
            [
                ['EO', '', '', '', '', ''],
                ['COA', '', '', '', '', ''],
                ['Processo SEI', '', '', '', '', ''],
                ['Ofício ANAC correlato', '', '', '', '', ''],
            ],
        )

    if 'bases e contatos operacionais' in n:
        return generic_registry_form(
            ['Base', 'Tipo', 'Contato principal', 'Contato contingência', 'Horário local', 'Observações'],
            [['', '', '', '', '', ''] for _ in range(8)],
        )

    if 'cadastro de tripulantes' in n:
        return generic_registry_form(
            ['Nome', 'CANAC', 'Função', 'Base', 'Vínculo', 'Status operacional', 'Contato'],
            [['', '', '', '', '', '☐ Apto  ☐ Restrito  ☐ Inativo', ''] for _ in range(8)],
        )

    if 'qualificação de tripulantes' in n:
        return generic_registry_form(
            ['Tripulante', 'Licença', 'CMA', 'Cheque RBAC 135', 'Treinamento recorrente', 'Validade final', 'Status'],
            [['', '', '', '', '', '', '☐ Válido  ☐ A vencer  ☐ Vencido'] for _ in range(8)],
        )

    if 'matriz de treinamento e elegibilidade' in n:
        return generic_registry_form(
            ['Treinando', 'Evento', 'Periodicidade', 'Janela', 'Data prevista', 'Data realizada', 'Elegível'],
            [['', '', '', '', '', '', '☐'] for _ in range(8)],
        )

    if 'provedores/instrutores/verificadores' in n:
        return generic_registry_form(
            ['Nome / empresa', 'Tipo', 'Escopo autorizado', 'Contrato até', 'Comprovação regulatória', 'Contato'],
            [['', '', '', '', '', ''] for _ in range(8)],
        )

    if 'dossiê de provedores e contratos de treinamento' in n:
        return generic_registry_form(
            ['Provedor', 'Contrato', 'Escopo', 'Validade', 'Auditoria / avaliação', 'Pendências'],
            [['', '', '', '', '', ''] for _ in range(6)],
        )

    if 'registro de treinamento individual' in n:
        return '\n\n'.join([
            generic_summary_fields([
                ['Tripulante', '____________________'],
                ['CANAC', '____________________'],
                ['Evento de treinamento', '____________________'],
                ['Provedor / instrutor', '____________________'],
                ['Período', '____/____/______ a ____/____/______'],
                ['Resultado', '☐ Aprovado  ☐ Reprovado  ☐ Pendente'],
            ]),
            make_table(
                ['Evidência', 'Identificação', 'Data', 'Arquivado'],
                [
                    ['Lista de presença', '', '', '☐'],
                    ['FAP / ficha de avaliação', '', '', '☐'],
                    ['Certificado', '', '', '☐'],
                    ['Prova teórica / prática', '', '', '☐'],
                ],
            ),
        ])

    if 'escala planejada e escala realizada' in n:
        return generic_registry_form(
            ['Tripulante', 'Data', 'Escala planejada', 'Escala realizada', 'Diferença justificada', 'Responsável'],
            [['', '', '', '', '☐', ''] for _ in range(10)],
        )

    if 'controle de jornada e descanso' in n:
        return generic_registry_form(
            ['Tripulante', 'Data', 'Jornada total', 'Tempo de voo', 'Sobreaviso', 'Descanso anterior', 'Conforme'],
            [['', '', '', '', '', '', '☐ Sim  ☐ Não'] for _ in range(10)],
        )

    if 'extrapolação de jornada' in n:
        return '\n\n'.join([
            generic_summary_fields([
                ['Tripulante', '____________________'],
                ['Data', '____/____/______'],
                ['Jornada prevista', '____________________'],
                ['Jornada realizada', '____________________'],
                ['Motivo da extrapolação', '____________________'],
                ['Comunicação à ANAC', '☐ Sim  ☐ Não'],
            ]),
            'Descrição objetiva:\n\n____________________________________________________________\n\n____________________________________________________________',
        ])

    if 'ocorrências de fadiga' in n or 'fadiga' in n:
        return '\n\n'.join([
            generic_summary_fields([
                ['Data / hora', '____/____/______ ________'],
                ['Tripulante', '____________________'],
                ['Missão / voo', '____________________'],
                ['Situação operacional', '____________________'],
            ]),
            '### Fatores contribuintes\n\n- ☐ Sono insuficiente\n- ☐ Jornada prolongada\n- ☐ Descanso inadequado\n- ☐ Estresse\n- ☐ Ambiente\n- ☐ Outro: ____________________',
            make_table(
                ['Medida', 'Aplicada'],
                [
                    ['Afastamento temporário', '☐'],
                    ['Replanejamento de escala', '☐'],
                    ['Avaliação adicional', '☐'],
                    ['Comunicação ao gestor', '☐'],
                ],
            ) + '\n\nDescrição:\n\n____________________________________________________________\n\n____________________________________________________________',
        ])

    if 'liberação/despacho' in n or 'despacho simplificado' in n:
        return '\n\n'.join([
            generic_summary_fields([
                ['Data', '____/____/______'],
                ['Missão / voo nº', '____________________'],
                ['Aeronave', '____________________'],
                ['PIC', '____________________'],
                ['Origem / destino', '____________________ / ____________________'],
            ]),
            make_table(
                ['Item de liberação', 'Sim', 'Não', 'Observações'],
                [
                    ['Tripulação habilitada', '☐', '☐', ''],
                    ['Documentação da aeronave válida', '☐', '☐', ''],
                    ['Meteorologia analisada', '☐', '☐', ''],
                    ['NOTAMs verificados', '☐', '☐', ''],
                    ['Combustível e reservas adequados', '☐', '☐', ''],
                    ['Peso e balanceamento dentro do envelope', '☐', '☐', ''],
                    ['Condição técnica apta', '☐', '☐', ''],
                ],
            ),
            'Decisão final: ☐ Liberado  ☐ Liberado com restrição  ☐ Não liberado\n\nRestrição / observação:\n\n____________________________________________________________',
        ])

    if 'decisões operacionais em missão' in n:
        return generic_registry_form(
            ['Data/hora', 'Missão', 'Decisão', 'Motivo', 'Autor da decisão', 'Comunicado a', 'Resultado'],
            [['', '', '', '', '', '', ''] for _ in range(8)],
        )

    if 'lista de passageiros' in n:
        return '\n\n'.join([
            generic_summary_fields([
                ['Data', '____/____/______'],
                ['Missão / voo nº', '____________________'],
                ['Aeronave', '____________________'],
                ['PIC', '____________________'],
                ['Origem / destino', '____________________ / ____________________'],
            ]),
            make_table(
                ['Nº', 'Nome completo', 'Documento', 'Contato emergência (nome)', 'Contato emergência (telefone)', 'Observações'],
                [[str(i), '', '', '', '', ''] for i in range(1, 7)],
            ),
            '### Declarações obrigatórias\n\n- ☐ Todos os passageiros receberam briefing de segurança.\n- ☐ Todos os passageiros informaram contato de emergência.\n- ☐ Uma via permanece a bordo e uma via permanece na base.\n\nAssinatura PIC: ____________________\n\nAssinatura base/despacho: ____________________',
        ])

    if 'cartão de instruções ao passageiro' in n:
        return '\n\n'.join([
            make_table(
                ['Item de briefing', 'Incluído', 'Observação de adaptação por aeronave'],
                [
                    ['Cinto de segurança', '☐', ''],
                    ['Portas e evacuação', '☐', ''],
                    ['Proibição de fumar e restrições', '☐', ''],
                    ['Uso de eletrônicos', '☐', ''],
                    ['Localização de extintor / primeiros socorros / sobrevivência', '☐', ''],
                    ['Procedimento em emergência', '☐', ''],
                ],
            ),
            make_table(
                ['Configuração da aeronave', 'Revisão do cartão', 'Vigência'],
                [['____________________', '____', '____/____/______']],
            ),
        ])

    if 'lista controlada de aeronaves' in n:
        return generic_registry_form(
            ['Matrícula', 'S/N', 'Modelo', 'Capacidade operacional', 'Base', 'Status', 'Observações'],
            [['', '', '', '', '', '☐ Ativa  ☐ Suspensa  ☐ Inativa', ''] for _ in range(8)],
        )

    if 'master control sheet de frota' in n:
        return generic_registry_form(
            ['Matrícula', 'Entrada na frota', 'Situação técnica', 'Situação documental', 'Próxima inspeção', 'Observações'],
            [['', '', '', '', '', ''] for _ in range(8)],
        )

    if 'dossiê de entrada/saída de aeronave' in n:
        return '\n\n'.join([
            generic_summary_fields([
                ['Matrícula', '____________________'],
                ['Tipo de processo', '☐ Entrada  ☐ Saída'],
                ['Data de avaliação', '____/____/______'],
                ['Responsável', '____________________'],
            ]),
            make_table(
                ['Documento / evidência', 'Presente', 'Referência'],
                [
                    ['Relatório técnico', '☐', ''],
                    ['Aprovação GR / DOPS', '☐', ''],
                    ['Contrato / distrato', '☐', ''],
                    ['Registros de aeronavegabilidade', '☐', ''],
                ],
            ),
        ])

    if 'configuração e aeronavegabilidade' in n:
        return generic_registry_form(
            ['Matrícula', 'Configuração cabine', 'Modificação/STC', 'Status aeronavegável', 'Data referência', 'Observações'],
            [['', '', '', '☐ Sim  ☐ Não', '', ''] for _ in range(8)],
        )

    if 'ficha de pesagem vigente' in n:
        return '\n\n'.join([
            generic_summary_fields([
                ['Matrícula', '____________________'],
                ['Data da pesagem', '____/____/______'],
                ['Peso vazio básico', '____________________'],
                ['Braço / momento de referência', '____________________'],
                ['Validade até', '____/____/______'],
            ]),
            'Observações técnicas:\n\n____________________________________________________________\n\n____________________________________________________________',
        ])

    if 'peso e balanceamento' in n:
        return '\n\n'.join([
            generic_summary_fields([
                ['Data', '____/____/______'],
                ['Missão / voo nº', '____________________'],
                ['Aeronave', '____________________'],
                ['PIC', '____________________'],
            ]),
            make_table(
                ['Item', 'Peso', 'Braço', 'Momento'],
                [
                    ['Peso vazio básico', '', '', ''],
                    ['Tripulação', '', '', ''],
                    ['Passageiros', '', '', ''],
                    ['Bagagem / carga', '', '', ''],
                    ['Combustível de decolagem', '', '', ''],
                    ['Total', '', '', ''],
                ],
            ),
            'Resultado no envelope AFM/POH: ☐ Dentro  ☐ Fora',
        ])

    if 'manifesto de carga' in n:
        return '\n\n'.join([
            generic_summary_fields([
                ['Data', '____/____/______'],
                ['Aeronave', '____________________'],
                ['PIC', '____________________'],
                ['Origem / destino', '____________________ / ____________________'],
            ]),
            make_table(
                ['Volume', 'Descrição', 'Peso (kg)', 'Posição', 'Fixação verificada', 'Observações'],
                [['', '', '', '', '☐', ''] for _ in range(6)],
            ),
        ])

    if 'biblioteca técnica afm/poh e suplementos' in n:
        return generic_registry_form(
            ['Documento técnico', 'Revisão', 'Aplicável à matrícula', 'Local disponível', 'Última conferência', 'Responsável'],
            [['', '', '', '', '', ''] for _ in range(8)],
        )

    if 'stcs e modificações' in n:
        return generic_registry_form(
            ['Matrícula', 'STC / modificação', 'Referência aprovação', 'Data incorporação', 'ICA associada', 'Status'],
            [['', '', '', '', '', '☐ Ativo  ☐ Removido'] for _ in range(8)],
        )

    if 'icas, sb e ad' in n:
        return generic_registry_form(
            ['Matrícula', 'ICA / SB / AD', 'Referência', 'Prazo / limite', 'Cumprimento em', 'OS / evidência', 'Status'],
            [['', '', '', '', '', '', '☐ Aberto  ☐ Cumprido  ☐ N/A'] for _ in range(10)],
        )

    if 'prestadores contratados' in n:
        return generic_registry_form(
            ['Prestador', 'Tipo', 'Escopo', 'Contrato até', 'Ponto focal', 'Contato', 'Status'],
            [['', '', '', '', '', '', '☐ Ativo  ☐ Suspenso'] for _ in range(8)],
        )

    if 'contratos de manutenção e slas' in n:
        return generic_registry_form(
            ['Contrato', 'Prestador', 'Escopo', 'SLA principal', 'Início', 'Fim', 'Status'],
            [['', '', '', '', '', '', '☐ Vigente  ☐ A vencer  ☐ Encerrado'] for _ in range(8)],
        )

    if 'matriz de inspeções periódicas' in n:
        return generic_registry_form(
            ['Matrícula', 'Inspeção', 'Critério (h/ciclos/calendário)', 'Limite', 'Próxima previsão', 'Responsável'],
            [['', '', '', '', '', ''] for _ in range(10)],
        )

    if 'plano de manutenção controlado' in n:
        return generic_registry_form(
            ['Item do plano', 'Referência técnica', 'Periodicidade', 'Próxima execução', 'Responsável', 'Observações'],
            [['', '', '', '', '', ''] for _ in range(10)],
        )

    if 'registros de manutenção e retorno ao serviço' in n:
        return generic_registry_form(
            ['Data', 'Matrícula', 'OS / workcard', 'Serviço executado', 'Responsável técnico', 'RTS emitido', 'Arquivo'],
            [['', '', '', '', '', '☐', ''] for _ in range(10)],
        )

    if 'formulários técnicos controlados' in n:
        return generic_registry_form(
            ['Formulário', 'Código', 'Revisão', 'Aplicação', 'Local oficial', 'Status'],
            [['', '', '', '', '', '☐ Vigente  ☐ Obsoleto'] for _ in range(8)],
        )

    if 'mal súbito/falecimento a bordo' in n:
        return '\n\n'.join([
            make_table(
                ['Etapa', 'Descrição controlada', 'Conferido'],
                [
                    ['1', 'Garantir segurança do voo e estabilizar a situação.', '☐'],
                    ['2', 'Aplicar procedimentos do PIC e coordenar comunicação ATS.', '☐'],
                    ['3', 'Acionar base / DOPS / apoio em solo.', '☐'],
                    ['4', 'Registrar passageiros envolvidos e atendimento prestado.', '☐'],
                    ['5', 'Abrir registro de ocorrência e preservar evidências.', '☐'],
                ],
            ),
            make_table(
                ['Contato crítico', 'Nome / órgão', 'Telefone / meio'],
                [['ATS / aeroporto', '', ''], ['Base operacional', '', ''], ['Autoridade local', '', '']],
            ),
        ])

    if 'programas específicos aplicáveis' in n or 'ppsp/psoa' in n:
        return generic_registry_form(
            ['Programa / declaração', 'Aplicável', 'Responsável', 'Revisão', 'Vigência', 'Evidência de implementação'],
            [['', '☐ Sim  ☐ Não', '', '', '', ''] for _ in range(8)],
        )

    if 'governança' in category:
        return generic_registry_form(
            ['Campo controlado', 'Valor / referência', 'Responsável', 'Data', 'Observações'],
            [['', '', '', '', ''] for _ in range(8)],
        )
    if 'opera' in category:
        return generic_registry_form(
            ['Campo operacional', 'Valor', 'Conferido', 'Responsável', 'Observações'],
            [['', '', '☐', '', ''] for _ in range(8)],
        )
    if 'aeronave' in category:
        return generic_registry_form(
            ['Matrícula', 'Campo técnico', 'Valor', 'Status', 'Observações'],
            [['', '', '', '', ''] for _ in range(8)],
        )
    if 'manutenção' in category:
        return generic_registry_form(
            ['Item', 'Referência', 'Prazo / vigência', 'Responsável', 'Observações'],
            [['', '', '', '', ''] for _ in range(8)],
        )

    return generic_registry_form(
        ['Campo', 'Preenchimento', 'Observações'],
        [['', '', ''] for _ in range(8)],
    )


def build_template(row: dict) -> str:
    parts = [
        f"# {row['documento']}",
        '',
        f"**Categoria:** {row['categoria']}  ",
        f"**Prioridade:** {row['prioridade']}  ",
        f"**Responsável sugerido:** {row['responsavel']}  ",
        f"**Fonte (manuais):** {row['fonte']}",
        '',
        section('Ficha de controle', common_control_table(row)),
        section('Finalidade e regra de uso', purpose_table(row)),
        section('Formulário controlado', form_body(row)),
        section('Validação obrigatória', validation_list()),
        section('Aprovações e ciência', approval_table()),
        section('Histórico de revisões', revision_table()),
    ]
    return '\n'.join(parts).strip() + '\n'


def main():
    markdown = CHECKLIST.read_text(encoding='utf-8')
    rows = parse_rows(markdown)
    if not rows:
        raise SystemExit('No rows parsed from checklist file.')

    OUT_DIR.mkdir(parents=True, exist_ok=True)

    index_lines = [
        '# Pacote de Documentos Externos (Formulários Controlados)',
        '',
        'Os arquivos abaixo foram estruturados como formulários semifechados, com campos fixos, tabelas, blocos de conferência e aprovações.',
        '',
        f'Total de formulários gerados: {len(rows)}',
        '',
        '| Documento | Arquivo |',
        '|---|---|',
    ]

    used_names = set()
    for idx, row in enumerate(rows, start=1):
        base = f"{idx:02d}_{slugify(row['documento'])}"
        name = base
        suffix = 2
        while name in used_names:
            name = f"{base}_{suffix}"
            suffix += 1
        used_names.add(name)

        file_name = f'{name}.md'
        path = OUT_DIR / file_name
        path.write_text(build_template(row), encoding='utf-8')
        index_lines.append(f"| {row['documento']} | [{file_name}](./{file_name}) |")

    (OUT_DIR / 'INDEX.md').write_text('\n'.join(index_lines) + '\n', encoding='utf-8')

    print(f'Generated {len(rows)} files in {OUT_DIR}')


if __name__ == '__main__':
    main()
