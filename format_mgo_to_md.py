from pathlib import Path
import re

src = Path('135/Asas MGO.txt')
dst = Path('135/Asas MGO.md')

SECTION_TITLES = {
    1: 'Prefácio',
    2: 'Estrutura Organizacional e Responsabilidades',
    3: 'Controle Operacional e Liberação de Voo',
    4: 'Qualificação e Treinamento',
    5: 'Procedimentos para Cumprimento da Jornada de Trabalho',
    6: 'Procedimentos com Passageiros',
    7: 'Transporte de Carga e Materiais de Características Especiais',
    8: 'Peso e Balanceamento',
    9: 'Combustível e Fluidos',
    10: 'Diário de Bordo',
    11: 'Procedimentos de Voo',
    12: 'Emergências',
    13: 'Anexos e Formulários Controlados',
}

SUBSECTION_MAP = [
    ('1.1 Finalidade do MGO', '1.1 Finalidade do MGO'),
    ('1.2 Escopo Operacional', '1.2 Escopo Operacional'),
    ('1.3 Referências Normativas', '1.3 Referências Normativas'),
    ('1.4 Organização do Sistema Documental', '1.4 Organização do Sistema Documental'),
    ('1.5 Lista de Detentores', '1.5 Lista de Detentores'),
    ('1.6 Sistema de Revisão', '1.6 Sistema de Revisão'),
    ('1.7 Registro de Revisões', '1.7 Registro de Revisões'),
    ('1.8 Lista de Páginas Efetivas', '1.8 Lista de Páginas Efetivas'),
    ('1.9 Abreviaturas e Acrônimos', '1.9 Abreviaturas e Acrônimos'),
    ('1.10 Definições', '1.10 Definições'),
    ('2.1 Princípios de Organização Operacional', '2.1 Princípios de Organização Operacional'),
    ('2.2 Gestor Responsável', '2.2 Gestor Responsável'),
    ('2.3 Diretor de Operações', '2.3 Diretor de Operações'),
    ('2.4 Tripulantes e Pessoal Operacional Designado', '2.4 Tripulantes e Pessoal Operacional Designado'),
    ('2.5 Delegação e Substituição Temporária', '2.5 Delegação e Substituição Temporária'),
    ('2.6 Gestão de Competência e Designação para Voo', '2.6 Gestão de Competência e Designação para Voo'),
    ('2.7 Interface com Manutenção e Segurança Operacional', '2.7 Interface com Manutenção e Segurança Operacional'),
    ('2.8 Registros, Rastreabilidade e Controle', '2.8 Registros, Rastreabilidade e Controle'),
    ('3.1 Princípios de Controle Operacional', '3.1 Princípios de Controle Operacional'),
    ('1.1 Pessoal Autorizado a Exercer Controle Operacional', '3.1.1 Pessoal Autorizado a Exercer Controle Operacional'),
    ('3.2 Planejamento e Validação Pré-voo', '3.2 Planejamento e Validação Pré-voo'),
    ('3.3 Métodos e Procedimentos para Iniciar, Continuar, Desviar e Terminar Voos', '3.3 Métodos e Procedimentos para Iniciar, Continuar, Desviar e Terminar Voos'),
    ('3.1 Iniciar', '3.3.1 Iniciar'),
    ('3.2 Continuar', '3.3.2 Continuar'),
    ('3.3 Desviar', '3.3.3 Desviar'),
    ('3.4 Terminar', '3.3.4 Terminar'),
    ('3.5 Liberação de Voo (Despacho Simplificado)', '3.3.5 Liberação de Voo (Despacho Simplificado)'),
    ('3.4 Suspensão, Cancelamento e Replanejamento', '3.4 Suspensão, Cancelamento e Replanejamento'),
    ('4.1 Supervisão de Voo e Protocolo de Aeronave Não Localizada', '3.4.1 Supervisão de Voo e Protocolo de Aeronave Não Localizada'),
    ('3.5 Gestão de Risco Operacional em Tempo Real', '3.5 Gestão de Risco Operacional em Tempo Real'),
    ('3.6 Controle de Jornada e Aptidão do Tripulante', '3.6 Controle de Jornada e Aptidão do Tripulante'),
    ('3.7 Registros de Controle Operacional', '3.7 Registros de Controle Operacional'),
    ('4.1 Princípios Gerais', '4.1 Princípios Gerais'),
    ('4.2 Requisitos de Qualificação do Piloto em Comando', '4.2 Requisitos de Qualificação do Piloto em Comando'),
    ('4.3 Verificações Obrigatórias — RBAC nº 135', '4.3 Verificações Obrigatórias — RBAC nº 135'),
    ('4.4 Treinamento Inicial e Periódico', '4.4 Treinamento Inicial e Periódico'),
    ('4.5 Instrutores e Verificadores', '4.5 Instrutores e Verificadores'),
    ('4.6 Controle Documental de Qualificações', '4.6 Controle Documental de Qualificações'),
    ('5.1 Política de Gestão de Fadiga', '5.1 Política de Gestão de Fadiga'),
    ('5.2 Limites Operacionais de Jornada (Lei nº 13.475/17)', '5.2 Limites Operacionais de Jornada (Lei nº 13.475/17)'),
    ('5.3 Escalas Planejadas e Realizadas', '5.3 Escalas Planejadas e Realizadas'),
    ('5.4 Responsabilidades do Tripulante para Gestão da Fadiga', '5.4 Responsabilidades do Tripulante para Gestão da Fadiga'),
    ('5.5 Comunicação de Extrapolação de Jornada', '5.5 Comunicação de Extrapolação de Jornada'),
    ('5.6 Procedimentos em Caso de Fadiga', '5.6 Procedimentos em Caso de Fadiga'),
    ('5.7 Registros e Retenção', '5.7 Registros e Retenção'),
    ('6.1 Princípios Gerais de Segurança com Passageiros', '6.1 Princípios Gerais de Segurança com Passageiros'),
    ('6.2 Procedimentos de Embarque e Desembarque', '6.2 Procedimentos de Embarque e Desembarque'),
    ('6.3 Transporte de Passageiros', '6.3 Transporte de Passageiros'),
    ('6.4 Lista de Passageiros', '6.4 Lista de Passageiros'),
    ('6.5 Briefing aos Passageiros', '6.5 Briefing aos Passageiros'),
    ('6.6 Requisitos de Segurança na Cabine', '6.6 Requisitos de Segurança na Cabine'),
    ('6.7 Procedimentos para Negar Embarque', '6.7 Procedimentos para Negar Embarque'),
    ('6.8 Transporte de Pessoas Fora das Provisões Regulares', '6.8 Transporte de Pessoas Fora das Provisões Regulares'),
    ('6.9 Operações Não Aplicáveis', '6.9 Operações Não Aplicáveis'),
    ('7.1 Transporte de Carga', '7.1 Transporte de Carga'),
    ('7.2 Materiais de Características Especiais', '7.2 Materiais de Características Especiais'),
    ('8.1 Princípios Gerais', '8.1 Princípios Gerais'),
    ('8.2 Responsabilidades', '8.2 Responsabilidades'),
    ('8.3 Procedimentos para Preenchimento do Manifesto de Carga', '8.3 Procedimentos para Preenchimento do Manifesto de Carga'),
    ('8.4 Pesos Padrão de Referência', '8.4 Pesos Padrão de Referência'),
    ('8.5 Configurações de Aeronave e Pesagem', '8.5 Configurações de Aeronave e Pesagem'),
    ('8.6 Registros e Retenção', '8.6 Registros e Retenção'),
    ('9.1 Princípios Gerais', '9.1 Princípios Gerais'),
    ('9.2 Responsabilidades', '9.2 Responsabilidades'),
    ('9.3 Planejamento de Combustível', '9.3 Planejamento de Combustível'),
    ('9.4 Abastecimento e Verificações Pré-voo', '9.4 Abastecimento e Verificações Pré-voo'),
    ('9.5 Óleo e Demais Fluidos', '9.5 Óleo e Demais Fluidos'),
    ('9.6 Registros e Retenção', '9.6 Registros e Retenção'),
    ('0.1 Princípios Gerais', '10.1 Princípios Gerais'),
    ('0.2 Responsabilidades', '10.2 Responsabilidades'),
    ('0.3 Preenchimento do Diário de Bordo', '10.3 Preenchimento do Diário de Bordo'),
    ('0.4 Correções e Integridade dos Registros', '10.4 Correções e Integridade dos Registros'),
    ('0.5 Registro de Discrepâncias Técnicas', '10.5 Registro de Discrepâncias Técnicas'),
    ('0.6 Guarda, Retenção e Disponibilidade', '10.6 Guarda, Retenção e Disponibilidade'),
    ('1.1 Princípios Gerais', '11.1 Princípios Gerais'),
    ('1.2 Responsabilidades', '11.2 Responsabilidades'),
    ('1.3 Planejamento Pré-voo', '11.3 Planejamento Pré-voo'),
    ('1.4 Plano de Voo ATS e Acompanhamento', '11.4 Plano de Voo ATS e Acompanhamento'),
    ('1.5 Condução do Voo e Tomada de Decisão em Rota', '11.5 Condução do Voo e Tomada de Decisão em Rota'),
    ('1.6 Documentação de Bordo e Uso de EFB', '11.6 Documentação de Bordo e Uso de EFB'),
    ('1.7 Procedimentos com Passageiros Durante o Voo', '11.7 Procedimentos com Passageiros Durante o Voo'),
    ('1.8 Procedimentos Pós-voo e Fechamento Operacional', '11.8 Procedimentos Pós-voo e Fechamento Operacional'),
    ('1.9 Registros Operacionais do Voo', '11.9 Registros Operacionais do Voo'),
    ('2.1 Princípios Gerais', '12.1 Princípios Gerais'),
    ('2.2 Autoridade do PIC em Emergência', '12.2 Autoridade do PIC em Emergência'),
    ('2.3 Ações Imediatas', '12.3 Ações Imediatas'),
    ('2.4 Coordenação com Órgãos ATS e Busca e Salvamento', '12.4 Coordenação com Órgãos ATS e Busca e Salvamento'),
    ('2.5 Pouso de Emergência e Assistência aos Ocupantes', '12.5 Pouso de Emergência e Assistência aos Ocupantes'),
    ('2.6 Notificações Mandatórias', '12.6 Notificações Mandatórias'),
    ('2.7 Registros e Preservação de Evidências', '12.7 Registros e Preservação de Evidências'),
    ('2.8 Treinamento e Preparação', '12.8 Treinamento e Preparação'),
    ('3.1 Princípios Gerais', '13.1 Princípios Gerais'),
    ('3.2 Lista de Anexos Controlados', '13.2 Lista de Anexos Controlados'),
    ('3.3 Controle de Revisão dos Anexos', '13.3 Controle de Revisão dos Anexos'),
    ('3.4 Disponibilidade e Uso Operacional', '13.4 Disponibilidade e Uso Operacional'),
    ('3.5 Modelos dos Anexos Operacionais', '13.5 Modelos dos Anexos Operacionais'),
]

INTROS = [
    'Este manual tem por finalidade:',
    'Este MGO aplica-se às operações de Asas de Socorro Táxi Aéreo com as seguintes premissas:',
    'Este MGO foi elaborado com base nos seguintes documentos (ou suas revisões vigentes):',
    'Para o perfil operacional adotado (Operador Simples), o sistema de documentos da empresa é composto, no mínimo, por:',
    'Para garantir estabilidade documental e reduzir revisões por mudanças administrativas, os seguintes dados variáveis devem permanecer fora do corpo do MGO, em documentos controlados externos:',
    'Compete ao Gestor Responsável, no âmbito da governança operacional:',
    'Compete ao Diretor de Operações:',
    'A designação de tripulantes para voo deve observar, cumulativamente:',
    'Antes de cada voo, devem ser confirmados, no mínimo:',
    'Constituem gatilhos mínimos para não liberação, suspensão ou cancelamento:',
    'A liberação deve ser formalizada por registro controlado (físico ou digital), com identificação de:',
    'Devem compor o acervo mínimo de rastreabilidade:',
    'Para ser designado como PIC em operações sob RBAC nº 135, o tripulante deve atender, no mínimo, aos seguintes requisitos:',
    'As verificações periódicas mandatórias para tripulantes em operação RBAC nº 135 incluem:',
    'O conteúdo dos treinamentos deve contemplar, no mínimo:',
    'O controle deve incluir, no mínimo:',
    'O PIC deve dispor as seguintes informações aos passageiros antes da decolagem:',
    'O PIC deve negar o embarque nas seguintes situações, sem prejuízo de outras condições de segurança:',
    'Conforme RBAC nº 135.85, a empresa pode transportar as seguintes pessoas fora das provisões de transporte de passageiros:',
    'As seguintes operações não são realizadas pela Asas de Socorro Táxi Aéreo e portanto não são objeto de procedimentos neste MGO:',
    'O manifesto deve conter, no mínimo:',
    'Compõem o conjunto mínimo de anexos operacionais deste MGO:',
]


def build_front_matter():
    return '\n'.join([
        '::: {custom-style="Title"}',
        'MGO — ASAS DE SOCORRO TÁXI AÉREO',
        ':::',
        '::: {custom-style="Subtitle"}',
        'Manual Geral de Operações',
        ':::',
        '::: {custom-style="Subtitle"}',
        'Operação RBAC 135 — Operador Simples',
        ':::',
        '',
        '```{=openxml}',
        '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
        '```',
        '',
        '### TERMO DE APROVAÇÃO',
        '',
        'Este Manual Geral de Operações (MGO) é elaborado em conformidade com os regulamentos e orientações vigentes da ANAC, em especial:',
        '',
        '- RBAC nº 119;',
        '- RBAC nº 135;',
        '- IS nº 119-004, Revisão M;',
        '- IS nº 135-002, revisão vigente;',
        '- IS e demais normativos aplicáveis ao escopo operacional autorizado nas Especificações Operativas (EO).',
        '',
        'Este MGO estabelece a organização geral da operação de Asas de Socorro Táxi Aéreo para operações não regulares, domésticas, com enquadramento de Operador Simples sob RBAC 135, em estrita observância às EO vigentes.',
        '',
        'A elaboração, implantação, atualização e controle deste MGO são de responsabilidade do Diretor de Operações da empresa, com governança documental definida no sistema de documentos operacionais.',
        '',
        'XXXXX Diretor de Operações — Asas de Socorro Táxi Aéreo',
        '',
        'XXXXX Gestor Responsável — Asas de Socorro Táxi Aéreo',
        '',
        'Nota de controle documental: os nomes dos ocupantes de função são mantidos em cadastro controlado externo, conforme política de dados variáveis deste manual.',
        '',
        'A Asas de Socorro Táxi Aéreo e cada pessoa a ela vinculada deve permitir, a qualquer tempo, que a ANAC realize inspeções ou exames — incluindo voo de acompanhamento — para verificar a conformidade com o Código Brasileiro de Aeronáutica, com os RBAC aplicáveis e com o COA e suas Especificações Operativas, conforme os artigos e parágrafos aplicáveis do CBA e do RBAC nº 119.',
        '',
        '### SUMÁRIO (VERSÃO INICIAL)',
        '',
    ])


def build_summary():
    return '\n'.join(f'- Seção {number} — {title}' for number, title in SECTION_TITLES.items())


def extract_body_seed():
    text = src.read_text(encoding='utf-8', errors='ignore').replace('\u00a0', ' ')
    text = re.sub(r'\s+', ' ', text).strip()
    body_marker = 'Seção 1 | PrefácioEste documento constitui o Manual Geral de Operações (MGO)'
    idx = text.find(body_marker)
    if idx != -1:
        return text[idx:]
    idx = text.rfind('Seção 1 | Prefácio')
    return text[idx:] if idx != -1 else text


def replace_heading(text, old, new, level=3):
    hashes = '#' * level
    pattern = re.compile(rf'(?:{re.escape(hashes)}\s*)?{re.escape(old)}')
    return pattern.sub(f'\n\n{hashes} {new}\n\n', text)


body = re.sub(r'\s+', ' ', extract_body_seed()).strip()

for number, title in SECTION_TITLES.items():
    body = replace_heading(body, f'Seção {number} | {title}', f'Seção {number} — {title}', level=2)

for old, new in SUBSECTION_MAP:
    body = replace_heading(body, old, new, level=3)

body = re.sub(r'Fim da Seção\s+(\d+)', r'\n\n**Fim da Seção \1**\n\n', body)
body = re.sub(r'Pendências para fechamento ANAC \(Seção\s+(\d+)\)', r'\n\n### Pendências para fechamento ANAC (Seção \1)\n\n', body)

for intro in INTROS:
    body = body.replace(intro, intro + '\n')

body = body.replace(
    'As seguintes funções estão autorizadas a exercer controle operacional na empresa:FunçãoGestor ResponsávelDiretor de OperaçõesOs nomes dos ocupantes dessas funções são mantidos em cadastro controlado externo, conforme Seção 1.4.',
    'As seguintes funções estão autorizadas a exercer controle operacional na empresa:\n- Gestor Responsável;\n- Diretor de Operações.\nOs nomes dos ocupantes dessas funções são mantidos em cadastro controlado externo, conforme Seção 1.4.'
)
body = body.replace(
    'Para a manutenção contratada junto a organização RBAC 145, a coordenação interna deve seguir a seguinte cadeia funcional:coordenador primário: Diretor de Operações (DOPS); coordenador alternativo: Gestor Responsável (GR), na indisponibilidade do DOPS.Compete ao coordenador interno de manutenção contratada monitorar vencimentos de inspeções e itens mandatórios, acionar formalmente a organização RBAC 145 para execução dos serviços requeridos e acompanhar o retorno documental de liberação.',
    'Para a manutenção contratada junto a organização RBAC 145, a coordenação interna deve seguir a seguinte cadeia funcional:\n- coordenador primário: Diretor de Operações (DOPS);\n- coordenador alternativo: Gestor Responsável (GR), na indisponibilidade do DOPS.\nCompete ao coordenador interno de manutenção contratada monitorar vencimentos de inspeções e itens mandatórios, acionar formalmente a organização RBAC 145 para execução dos serviços requeridos e acompanhar o retorno documental de liberação.'
)
body = re.sub(r'(?<=[A-Za-zÀ-ÿ0-9\)])\.(?=[A-ZÁÉÍÓÚÂÊÔÃÕ])', '.\n', body)
body = re.sub(r';(?=[A-Za-zÁÉÍÓÚÂÊÔÃÕ0-9\(])', ';\n', body)
body = re.sub(r'(?<=[a-zà-ÿ0-9\)])\. (?=[A-ZÁÉÍÓÚÂÊÔÃÕ])', '.\n', body)
body = re.sub(r'(\d+)\.\s+(\d)\s+(\d)(?=\D|$)', r'\1.\2\3', body)
body = re.sub(r'(RBAC nº\s+\d+)\.\s+(\d+)', r'\1.\2', body)
body = re.sub(r'(?<=\.)\s*[3479]\.($|\n)', r'\1', body)
body = re.sub(r'\n\s*[0-9]+\.?\s*\n', '\n', body)
body = re.sub(r'\n{3,}', '\n\n', body)

bullet_prefix_re = re.compile(
    r'^(RBAC nº|IS nº|AFM/POH|demais IS|MGO \(este manual\)|SOP vigente|PTO vigente|MGM vigente|EO vigentes|listas e cadastros|formulários e instruções|nomes dos ocupantes|cadastro de tripulantes|lista de aeronaves|contatos operacionais|ANAC \(|Gestor Responsável;|Diretor de Operações;|responsável\(is\)|cópia controlada|acervo digital|inserir código interno|substituir campos|atualizar referência|as definições do RBAC|as definições constantes|estabelecer a estrutura|consolidar, em nível|definir a integração|assegurar rastreabilidade|operação doméstica|operação sob RBAC|base principal:|operação conduzida|operação sem escopos|assegurar que a empresa|garantir que políticas|manter a estrutura|assegurar que não haja|apoiar o processo|zelar pelo cumprimento|assegurar a manutenção|implementar, manter|assegurar que designações|coordenar a atualização|assegurar aderência|exercer autoridade|garantir rastreabilidade|colocar à disposição|designar as pessoas|cadastro vigente|requisitos de treinamento|limitações operacionais|conformidade da missão|aptidão legal|situação de aeronavegabilidade|avaliação meteorológica|adequação de combustível|disponibilidade de documentação|voo/missão|tripulante designado|aeronave designada|status de conformidade|indisponibilidade técnica|condição meteorológica|perda de requisito|indisponibilidade documental|o GR e o DOPS|é iniciado o protocolo|a ANAC é notificada|registros de liberação|evidências de validação|registros de discrepâncias|registros de decisões|habilitação de voo válida|qualificação e familiarização|cumprimento dos requisitos|verificação de proficiência|certificado médico válido|experiência recente|treinamento de emergência|conhecimento de equipamentos|procedimentos normais|procedimentos de cabine|segurança operacional|demais matérias exigidas|habilitação e validade|qualificação de tipo|registros de treinamento|vencimentos de verificações|proibição de fumar|proibição de consumo|proibição de transporte|uso de cintos|ajuste dos encostos|localização e instruções|restrições ao uso|localização e conteúdo|passageiro que|colaborador em voo|pessoa exercendo|servidor designado|pessoa autorizada|operações aeromédicas|operações de ligação|transporte de artigos|operações internacionais|matrícula da aeronave|identificação do PIC|origem e destino|número de passageiros|peso e alocação|quantidade de combustível|peso vazio básico|peso total de rampa|Anexo 13-A|Anexo 13-B|Anexo 13-C|Anexo 13-D|Anexo 13-E|Anexo 13-F)'
)

lines = []
for raw in body.split('\n'):
    line = raw.strip()
    if not line:
        if lines and lines[-1] != '':
            lines.append('')
        continue
    if line.startswith('## ') or line.startswith('### ') or line.startswith('**Fim da Seção'):
        if lines and lines[-1] != '':
            lines.append('')
        lines.append(line)
        lines.append('')
        continue
    if re.fullmatch(r'(\*\*\s*)?[#*]+', line):
        continue
    if bullet_prefix_re.match(line) and not line.startswith('- '):
        lines.append('- ' + line)
        continue
    lines.append(line)

normalized = []
for line in lines:
    stripped = line.strip()
    if stripped.startswith('- ') and ';' in stripped:
        parts = [part.strip().rstrip(';') for part in stripped[2:].split(';') if part.strip()]
        for index, part in enumerate(parts):
            suffix = ';' if index < len(parts) - 1 else ''
            normalized.append(f'- {part}{suffix}')
        continue
    if stripped.startswith('confirmar ') and ';' in stripped:
        parts = [part.strip().rstrip(';') for part in stripped.split(';') if part.strip()]
        normalized.extend([f'- {part}' for part in parts])
        continue
    if stripped in {'4.', '3.', '7.', '9.'}:
        continue
    normalized.append(line)

clean = []
prev_blank = False
for line in normalized:
    blank = line == ''
    if blank and prev_blank:
        continue
    clean.append(line)
    prev_blank = blank

text_out = build_front_matter() + build_summary() + '\n\n' + '\n'.join(clean).strip() + '\n'
dst.write_text(text_out, encoding='utf-8')
print(dst)
