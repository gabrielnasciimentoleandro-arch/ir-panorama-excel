#!/usr/bin/env python3
"""Gera a planilha autoral IR Panorama para o desafio da DIO.

O arquivo publicado usa somente dados fictícios. Não há macros, APIs,
conexões externas ou envio de dados pela internet.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

from openpyxl import Workbook
from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "ir_panorama_organizador.xlsx"
AUTHOR = "0Barone (@gabrielnasciimentoleandro-arch)"
AUTHOR_HANDLE = "gabrielnasciimentoleandro-arch"

# Identidade visual autoral: painel fiscal com alto contraste.
NAVY = "0F172A"
NAVY_2 = "1E293B"
SLATE = "334155"
CYAN = "38BDF8"
CYAN_DARK = "0284C7"
TEAL = "14B8A6"
EMERALD = "10B981"
AMBER = "F59E0B"
RED = "EF4444"
WHITE = "FFFFFF"
TEXT = "172033"
MUTED = "64748B"
LIGHT = "F8FAFC"
INPUT = "ECFEFF"
CALC = "E2E8F0"
PALE_GREEN = "DCFCE7"
PALE_AMBER = "FEF3C7"
PALE_RED = "FEE2E2"
BORDER = "CBD5E1"

thin = Side(style="thin", color=BORDER)
medium_cyan = Side(style="medium", color=CYAN)

CURRENCY = 'R$ #,##0.00;[Red]-R$ #,##0.00'
PERCENT = '0%'
DATE_FMT = 'dd/mm/yyyy'
INTEGER = '0'

YEAR = 2025
EXERCISE = 2026
DATA_FIRST_ROW = 9
DATA_LAST_ROW = 38

NAV_LINKS = [
    ("Início", "Início"),
    ("Titular", "Titular"),
    ("Dependentes", "Dependentes"),
    ("Rendimentos", "Rendimentos"),
    ("Bens", "Bens"),
    ("Pagamentos", "Pagamentos"),
    ("Informes", "Informes"),
    ("Notas", "Notas"),
    ("Guia", "Guia"),
]

SUPPORT_LISTS: dict[str, list[str]] = {
    "lista_status_documento": ["Pendente", "Recebido", "Conferido", "Não se aplica"],
    "lista_categorias_rendimento": [
        "Trabalho assalariado", "Trabalho autônomo", "Aposentadoria/Pensão",
        "Rendimentos isentos", "Tributação exclusiva", "Aluguéis",
        "Exterior", "Outros",
    ],
    "lista_grupos_bens": [
        "Imóveis", "Veículos", "Contas e depósitos", "Aplicações e investimentos",
        "Participações societárias", "Criptoativos", "Créditos", "Outros bens",
    ],
    "lista_categorias_pagamento": [
        "Plano de saúde", "Médico/Hospital", "Dentista", "Educação",
        "Previdência", "Pensão alimentícia", "Profissional liberal", "Outros",
    ],
    "lista_prioridades": ["Baixa", "Média", "Alta", "Crítica"],
    "lista_status_nota": ["Aberta", "Em andamento", "Concluída"],
    "lista_estados_civis": [
        "Solteiro(a)", "Casado(a)", "União estável", "Separado(a)",
        "Divorciado(a)", "Viúvo(a)",
    ],
    "lista_ufs": [
        "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT",
        "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO",
        "RR", "SC", "SP", "SE", "TO",
    ],
    "lista_relacoes_dependente": [
        "Cônjuge/Companheiro(a)", "Filho(a)", "Enteado(a)", "Pai/Mãe",
        "Avô/Avó", "Irmão(ã)", "Neto(a)", "Outro permitido",
    ],
    "lista_sim_nao": ["Sim", "Não"],
    "lista_titular_dependente": ["Titular", "Dependente"],
    "lista_categorias_informe": [
        "Rendimentos", "Bancos", "Investimentos", "Saúde", "Educação",
        "Previdência", "Bens e direitos", "Dívidas", "Outros",
    ],
    "lista_tipos_declaracao": ["Original", "Retificadora"],
    "lista_status_cadastro": ["Pendente", "Revisar", "Completo"],
    "lista_categorias_nota": [
        "Titular", "Dependente", "Rendimento", "Bem", "Pagamento",
        "Informe", "Conferência final", "Outro",
    ],
    "lista_paises": ["Brasil", "Estados Unidos", "Portugal", "Outro"],
}

RENDIMENTOS = [
    ["REN-001", "Trabalho assalariado", "Empresa Horizonte Ltda.", "00.000.000/0001-00", 84000.00, 7800.00, 6300.00, "informe_empresa_2025.pdf", "Conferido", "Exemplo fictício"],
    ["REN-002", "Rendimentos isentos", "Banco Aurora S.A.", "00.000.000/0002-00", 1650.35, 0.00, 0.00, "informe_banco_2025.pdf", "Conferido", "Poupança — exemplo"],
    ["REN-003", "Tributação exclusiva", "Corretora Modelo S.A.", "00.000.000/0003-00", 980.40, 0.00, 0.00, "informe_corretora_2025.pdf", "Recebido", "Conferir classificação"],
    ["REN-004", "Trabalho autônomo", "Clientes diversos", "N/A", 7200.00, 0.00, 960.00, "livro_caixa_exemplo.xlsx", "Pendente", "Consolidar recibos"],
]

BENS = [
    ["BEM-001", "Contas e depósitos", "06.01", "Conta corrente fictícia", "Brasil", date(2021, 3, 10), 4200.00, 6850.25, "extrato_conta_2025.pdf", "Conferido", "Exemplo fictício"],
    ["BEM-002", "Aplicações e investimentos", "04.02", "Aplicação de renda fixa fictícia", "Brasil", date(2022, 8, 15), 21000.00, 27800.50, "informe_investimento_2025.pdf", "Recebido", "Conferir rendimentos"],
    ["BEM-003", "Veículos", "02.01", "Veículo de demonstração", "Brasil", date(2020, 6, 20), 58000.00, 58000.00, "documento_veiculo.pdf", "Conferido", "Custo histórico"],
    ["BEM-004", "Imóveis", "01.12", "Apartamento fictício", "Brasil", date(2019, 2, 11), 245000.00, 245000.00, "contrato_imovel_exemplo.pdf", "Conferido", "Custo histórico"],
]

PAGAMENTOS = [
    ["PAG-001", "Plano de saúde", "Saúde Modelo S.A.", "00.000.000/0004-00", "Titular", 4560.00, 0.00, None, "informe_saude_2025.pdf", "Conferido", "Exemplo fictício"],
    ["PAG-002", "Dentista", "Clínica Exemplo", "00.000.000/0005-00", "Titular", 1800.00, 300.00, None, "recibos_dentista.pdf", "Recebido", "Conferir reembolso"],
    ["PAG-003", "Educação", "Escola Demonstração", "00.000.000/0006-00", "Dependente", 6400.00, 0.00, None, "comprovante_escolar.pdf", "Pendente", "Aguardando informe anual"],
]

DEPENDENTES = [
    ["DEP-001", "Lucas Exemplo", "000.000.000-00", date(2014, 5, 12), "Filho(a)", "Sim", "Revisar", "Dados inteiramente fictícios"],
]

INFORMES = [
    ["INF-001", "Rendimentos", "Empresa Horizonte Ltda.", "00.000.000/0001-00", "Informe de rendimentos", "docs/informe_empresa_2025.pdf", date(2026, 2, 10), "Conferido", "Exemplo fictício"],
    ["INF-002", "Bancos", "Banco Aurora S.A.", "00.000.000/0002-00", "Informe bancário", "docs/informe_banco_2025.pdf", date(2026, 2, 12), "Conferido", "Exemplo fictício"],
    ["INF-003", "Investimentos", "Corretora Modelo S.A.", "00.000.000/0003-00", "Informe da corretora", "docs/informe_corretora_2025.pdf", date(2026, 2, 18), "Recebido", "Conferir classificação"],
    ["INF-004", "Saúde", "Saúde Modelo S.A.", "00.000.000/0004-00", "Informe do plano", "docs/informe_saude_2025.pdf", date(2026, 2, 15), "Conferido", "Exemplo fictício"],
    ["INF-005", "Saúde", "Clínica Exemplo", "00.000.000/0005-00", "Recibos e reembolsos", "docs/recibos_dentista.pdf", date(2026, 2, 20), "Recebido", "Conferir reembolso"],
    ["INF-006", "Educação", "Escola Demonstração", "00.000.000/0006-00", "Comprovante anual", "docs/comprovante_escolar.pdf", None, "Pendente", "Solicitar à instituição"],
    ["INF-007", "Previdência", "INSS — exemplo", "N/A", "Extrato previdenciário", "docs/extrato_inss.pdf", date(2026, 2, 14), "Conferido", "Exemplo fictício"],
    ["INF-008", "Bens e direitos", "Órgão de trânsito — exemplo", "N/A", "Documento do veículo", "docs/documento_veiculo.pdf", date(2026, 1, 28), "Conferido", "Exemplo fictício"],
]

NOTAS = [
    ["NOT-001", "Rendimento", "Conferir informe da corretora", "Validar a classificação antes da revisão final.", "Alta", date(2026, 4, 15), "Em andamento", None, "Titular"],
    ["NOT-002", "Informe", "Solicitar comprovante escolar", "Pedir o documento anual à instituição.", "Alta", date(2026, 4, 20), "Aberta", None, "Titular"],
    ["NOT-003", "Titular", "Revisar dados cadastrais", "Confirmar endereço e telefone da cópia privada.", "Média", date(2026, 4, 25), "Concluída", None, "Titular"],
]


def fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


def border(accent: bool = False) -> Border:
    return Border(
        left=medium_cyan if accent else thin,
        right=thin,
        top=thin,
        bottom=thin,
    )


def merge_box(ws, cell_range: str, value, *, bg: str = WHITE, color: str = TEXT,
              size: int = 11, bold: bool = False, align: str = "center",
              number_format: str | None = None, hyperlink: str | None = None) -> None:
    ws.merge_cells(cell_range)
    top_left = cell_range.split(":")[0]
    cell = ws[top_left]
    cell.value = value
    cell.fill = fill(bg)
    cell.font = Font(name="Aptos", size=size, bold=bold, color=color)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = border()
    if number_format:
        cell.number_format = number_format
    if hyperlink:
        cell.hyperlink = hyperlink


def add_internal_link(cell, sheet_name: str) -> None:
    cell.hyperlink = f"#'{sheet_name}'!A1"


def setup_sheet(ws, title: str, subtitle: str, *, freeze: str | None = None,
                zoom: int = 90) -> None:
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = zoom
    ws.freeze_panes = freeze
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.4
    ws.page_margins.bottom = 0.4
    ws.oddFooter.center.text = f"IR Panorama • {AUTHOR}"
    ws.oddFooter.center.size = 8
    ws.oddFooter.center.color = MUTED

    for col in range(1, 13):
        ws.column_dimensions[chr(64 + col)].width = 14

    for row in range(1, 4):
        ws.row_dimensions[row].height = 24
        for col in range(1, 13):
            ws.cell(row, col).fill = fill(NAVY)

    ws.merge_cells("A1:H1")
    ws["A1"] = title
    ws["A1"].font = Font(name="Aptos Display", size=20, bold=True, color=WHITE)
    ws["A1"].alignment = Alignment(vertical="center")
    ws.merge_cells("I1:L1")
    ws["I1"] = "PROJETO AUTORAL • DADOS FICTÍCIOS"
    ws["I1"].font = Font(name="Aptos", size=8, bold=True, color=CYAN)
    ws["I1"].alignment = Alignment(horizontal="right", vertical="center")

    for col, (label, target) in enumerate(NAV_LINKS, 1):
        c = ws.cell(2, col, label.upper())
        c.fill = fill(CYAN_DARK if label == ws.title else NAVY_2)
        c.font = Font(name="Aptos", size=7.5, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(bottom=medium_cyan)
        add_internal_link(c, target)
    for col in range(10, 13):
        ws.cell(2, col).fill = fill(NAVY_2)
        ws.cell(2, col).border = Border(bottom=medium_cyan)

    ws.merge_cells("A3:L3")
    ws["A3"] = subtitle
    ws["A3"].fill = fill(LIGHT)
    ws["A3"].font = Font(name="Aptos", size=9, italic=True, color=MUTED)
    ws["A3"].alignment = Alignment(vertical="center")
    ws["A3"].border = Border(bottom=thin)


def style_section_header(cell, text: str) -> None:
    cell.value = text
    cell.fill = fill(NAVY_2)
    cell.font = Font(name="Aptos", size=10, bold=True, color=WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = border(accent=True)


def style_input(cell, number_format: str | None = None, align: str = "left") -> None:
    cell.fill = fill(INPUT)
    cell.font = Font(name="Aptos", size=10, color=TEXT)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = border(accent=True)
    cell.protection = Protection(locked=False)
    if number_format:
        cell.number_format = number_format


def style_formula(cell, number_format: str | None = None, align: str = "center") -> None:
    cell.fill = fill(CALC)
    cell.font = Font(name="Aptos", size=10, bold=True, color=TEXT)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = border()
    if number_format:
        cell.number_format = number_format


def add_defined_name(wb: Workbook, name: str, reference: str) -> None:
    wb.defined_names.add(DefinedName(name=name, attr_text=reference))


def add_list_validation(ws, target: str, named_range: str, prompt: str) -> None:
    dv = DataValidation(type="list", formula1=f"={named_range}", allow_blank=True)
    dv.errorTitle = "Valor inválido"
    dv.error = "Escolha uma opção da lista."
    dv.promptTitle = "Lista de opções"
    dv.prompt = prompt
    dv.showErrorMessage = True
    dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(target)


def add_decimal_validation(ws, target: str) -> None:
    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    dv.errorTitle = "Valor inválido"
    dv.error = "Digite um valor igual ou maior que zero."
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(target)


def add_date_validation(ws, target: str, start_year: int = 1900, end_year: int = 2100) -> None:
    dv = DataValidation(
        type="date", operator="between",
        formula1=f"DATE({start_year},1,1)", formula2=f"DATE({end_year},12,31)",
        allow_blank=True,
    )
    dv.errorTitle = "Data inválida"
    dv.error = f"Informe uma data entre {start_year} e {end_year}."
    dv.showErrorMessage = True
    ws.add_data_validation(dv)
    dv.add(target)


def card(ws, start_col: int, end_col: int, label_row: int, value_start: int,
         value_end: int, label: str, formula: str, fmt: str, accent: str) -> str:
    start_letter = chr(64 + start_col)
    end_letter = chr(64 + end_col)
    merge_box(ws, f"{start_letter}{label_row}:{end_letter}{label_row}", label,
              bg=NAVY_2, color=WHITE, size=8, bold=True)
    merge_box(ws, f"{start_letter}{value_start}:{end_letter}{value_end}", formula,
              bg=WHITE, color=accent, size=18, bold=True, number_format=fmt)
    return f"{start_letter}{value_start}"


def nav_card(ws, cell_range: str, number: str, label: str, description: str,
             target: str, accent: str) -> None:
    text = f"{number}  {label}\n{description}"
    merge_box(ws, cell_range, text, bg=WHITE, color=TEXT, size=10, bold=True,
              align="left", hyperlink=f"#'{target}'!A1")
    anchor = ws[cell_range.split(":")[0]]
    anchor.border = Border(left=Side(style="thick", color=accent), right=thin, top=thin, bottom=thin)


def build_dashboard(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "Início"
    setup_sheet(
        ws,
        "IR PANORAMA",
        "Organizador autoral para reunir documentos, dados e pendências da declaração de Imposto de Renda.",
        freeze="A4",
        zoom=85,
    )
    widths = {
        "A": 14, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14,
        "G": 14, "H": 14, "I": 14, "J": 14, "K": 14, "L": 14,
        "N": 18, "O": 12,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for row in range(4, 52):
        ws.row_dimensions[row].height = 22

    ws.merge_cells("A5:L5")
    ws["A5"] = "VISÃO GERAL DA ORGANIZAÇÃO"
    style_section_header(ws["A5"], "VISÃO GERAL DA ORGANIZAÇÃO")
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center")

    merge_box(ws, "A6:B7", "ANO-CALENDÁRIO", bg=NAVY_2, color=WHITE, size=9, bold=True)
    merge_box(ws, "C6:D7", YEAR, bg=INPUT, color=CYAN_DARK, size=18, bold=True,
              number_format=INTEGER)
    ws["C6"].protection = Protection(locked=False)
    merge_box(ws, "E6:F7", "EXERCÍCIO", bg=NAVY_2, color=WHITE, size=9, bold=True)
    merge_box(ws, "G6:H7", "=ano_calendario+1", bg=CALC, color=EMERALD, size=18,
              bold=True, number_format=INTEGER)
    merge_box(ws, "I6:J7", "MODO", bg=NAVY_2, color=WHITE, size=9, bold=True)
    merge_box(ws, "K6:L7", "DEMONSTRAÇÃO", bg=PALE_AMBER, color=TEXT, size=11, bold=True)
    ws["C6"].comment = Comment("Altere para o ano dos documentos que serão organizados.", "IR Panorama")

    card(ws, 1, 3, 9, 10, 12, "RENDIMENTOS BRUTOS",
         "=SUM(Rendimentos!$E$9:$E$38)", CURRENCY, CYAN_DARK)
    card(ws, 4, 6, 9, 10, 12, "BENS — SITUAÇÃO ATUAL",
         "=SUM(Bens!$H$9:$H$38)", CURRENCY, EMERALD)
    card(ws, 7, 9, 9, 10, 12, "PAGAMENTOS LÍQUIDOS",
         "=SUM(Pagamentos!$H$9:$H$38)", CURRENCY, AMBER)
    card(ws, 10, 12, 9, 10, 12, "DOCUMENTOS CONFERIDOS",
         '=IFERROR((COUNTIF(Informes!$H$9:$H$38,"Conferido")+COUNTIF(Informes!$H$9:$H$38,"Não se aplica"))/COUNTIF(Informes!$A$9:$A$38,"<>"),0)',
         PERCENT, TEAL)

    card(ws, 1, 3, 14, 15, 16, "IR RETIDO INFORMADO",
         "=SUM(Rendimentos!$F$9:$F$38)", CURRENCY, CYAN_DARK)
    card(ws, 4, 6, 14, 15, 16, "INSS INFORMADO",
         "=SUM(Rendimentos!$G$9:$G$38)", CURRENCY, EMERALD)
    card(ws, 7, 9, 14, 15, 16, "PENDÊNCIAS ABERTAS",
         '=COUNTIFS(Informes!$A$9:$A$38,"<>",Informes!$H$9:$H$38,"<>Conferido",Informes!$H$9:$H$38,"<>Não se aplica")+COUNTIFS(Notas!$A$9:$A$38,"<>",Notas!$G$9:$G$38,"<>Concluída")',
         INTEGER, RED)
    card(ws, 10, 12, 14, 15, 16, "REGISTROS FINANCEIROS",
         "=COUNTA(Rendimentos!$A$9:$A$38)+COUNTA(Bens!$A$9:$A$38)+COUNTA(Pagamentos!$A$9:$A$38)",
         INTEGER, SLATE)

    merge_box(ws, "A18:L18", "PROGRESSO DA CONFERÊNCIA DOCUMENTAL", bg=NAVY_2,
              color=WHITE, size=9, bold=True)
    merge_box(ws, "A19:L20",
              '=IFERROR((COUNTIF(Informes!$H$9:$H$38,"Conferido")+COUNTIF(Informes!$H$9:$H$38,"Não se aplica"))/COUNTIF(Informes!$A$9:$A$38,"<>"),0)',
              bg=LIGHT, color=CYAN_DARK, size=18, bold=True, number_format=PERCENT)
    ws.conditional_formatting.add(
        "A19:L20",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                    color=TEAL, showValue=True),
    )

    merge_box(ws, "A22:L22", "MENU PRINCIPAL — CLIQUE EM UMA ÁREA", bg=NAVY,
              color=WHITE, size=10, bold=True)
    nav_card(ws, "A23:C26", "01", "TITULAR", "Cadastro e parâmetros da declaração", "Titular", CYAN)
    nav_card(ws, "D23:F26", "02", "DEPENDENTES", "Relações e conferência cadastral", "Dependentes", TEAL)
    nav_card(ws, "G23:I26", "03", "RENDIMENTOS", "Fontes, IRRF e INSS", "Rendimentos", EMERALD)
    nav_card(ws, "J23:L26", "04", "BENS", "Saldos anterior e atual", "Bens", AMBER)
    nav_card(ws, "A28:C31", "05", "PAGAMENTOS", "Valores, reembolsos e documentos", "Pagamentos", AMBER)
    nav_card(ws, "D28:F31", "06", "INFORMES", "Checklist de documentos", "Informes", TEAL)
    nav_card(ws, "G28:I31", "07", "NOTAS", "Prazos, prioridades e alertas", "Notas", RED)
    nav_card(ws, "J28:L31", "08", "GUIA", "Uso, privacidade e validação", "Guia", CYAN)

    merge_box(ws, "A33:F33", "RESUMO DE REVISÃO", bg=NAVY_2, color=WHITE, size=9, bold=True)
    summary = [
        (35, "Cadastro do titular", '=IF(COUNTA(Titular!$C$8:$C$20)>=10,"Pronto","Revisar")'),
        (37, "Informes", '=COUNTIF(Informes!$H$9:$H$38,"Conferido")&" de "&COUNTIF(Informes!$A$9:$A$38,"<>")&" conferidos"'),
        (39, "Notas", '=COUNTIFS(Notas!$A$9:$A$38,"<>",Notas!$G$9:$G$38,"<>Concluída")&" notas abertas"'),
        (41, "Documentos", '=IF(COUNTIF(Informes!$H$9:$H$38,"Pendente")=0,"Sem documentos pendentes",COUNTIF(Informes!$H$9:$H$38,"Pendente")&" documento(s) pendente(s)")'),
        (43, "Próxima ação", '=IF(G15=0,"Pronto para revisão final","Revise as pendências")'),
    ]
    for row, label, formula in summary:
        merge_box(ws, f"A{row}:B{row}", label, bg=LIGHT, color=MUTED, size=9, bold=True,
                  align="left")
        merge_box(ws, f"C{row}:F{row}", formula, bg=WHITE, color=TEXT, size=10, bold=True,
                  align="left")

    # Dados auxiliares do gráfico (fora da área de impressão).
    ws["N2"] = "Status"
    ws["O2"] = "Quantidade"
    statuses = ["Conferido", "Recebido", "Pendente", "Não se aplica"]
    for row, status in enumerate(statuses, 3):
        ws.cell(row, 14, status)
        ws.cell(row, 15, f'=COUNTIF(Informes!$H$9:$H$38,"{status}")')
    ws.column_dimensions["N"].hidden = True
    ws.column_dimensions["O"].hidden = True

    chart = DoughnutChart()
    chart.title = "Situação dos informes"
    chart.style = 10
    chart.holeSize = 58
    chart.height = 7.5
    chart.width = 13
    labels = Reference(ws, min_col=14, min_row=3, max_row=6)
    data = Reference(ws, min_col=15, min_row=2, max_row=6)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showPercent = True
    chart.dataLabels.showLeaderLines = True
    chart.legend.position = "b"
    ws.add_chart(chart, "G33")

    merge_box(ws, "A46:L46", "LINKS OFICIAIS — ABREM NO NAVEGADOR", bg=NAVY_2,
              color=WHITE, size=9, bold=True)
    merge_box(ws, "A47:D48", "MEU IMPOSTO DE RENDA", bg=WHITE, color=CYAN_DARK,
              size=10, bold=True, hyperlink="https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda")
    merge_box(ws, "E47:H48", "PORTAL e-CAC", bg=WHITE, color=CYAN_DARK,
              size=10, bold=True, hyperlink="https://cav.receita.fazenda.gov.br/autenticacao/login")
    merge_box(ws, "I47:L48", "PORTAL GOV.BR", bg=WHITE, color=CYAN_DARK,
              size=10, bold=True, hyperlink="https://www.gov.br/pt-br")

    merge_box(ws, "A50:L51",
              f"Autoria: {AUTHOR} • Arquivo demonstrativo com dados fictícios • Não substitui o programa oficial nem orientação profissional.",
              bg=NAVY, color=WHITE, size=9, bold=True)
    ws.print_area = "A1:L51"


def build_titular(wb: Workbook) -> None:
    ws = wb.create_sheet("Titular")
    setup_sheet(ws, "TITULAR", "Cadastre os dados básicos na cópia privada. Os valores deste arquivo são fictícios.", freeze="A7")
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 27
    ws.column_dimensions["C"].width = 34
    ws.column_dimensions["D"].width = 4
    ws.column_dimensions["E"].width = 5
    ws.column_dimensions["F"].width = 24
    ws.column_dimensions["G"].width = 4
    ws.column_dimensions["H"].width = 26
    ws.column_dimensions["I"].width = 22
    ws.column_dimensions["J"].width = 12
    ws.column_dimensions["K"].width = 12
    ws.column_dimensions["L"].width = 12

    merge_box(ws, "B5:C5", "DADOS DO TITULAR", bg=NAVY_2, color=WHITE, size=10, bold=True)
    labels_values = [
        (8, "Nome completo", "Marina Exemplo"),
        (9, "CPF", "000.000.000-00"),
        (10, "Data de nascimento", date(1990, 1, 15)),
        (11, "E-mail", "marina.exemplo@exemplo.invalid"),
        (12, "Telefone", "(00) 00000-0000"),
        (13, "Ocupação", "Analista de projetos"),
        (14, "Estado civil", "Solteiro(a)"),
        (15, "Endereço", "Endereço fictício para demonstração"),
        (16, "Cidade", "Cidade Exemplo"),
        (17, "UF", "SP"),
        (18, "CEP", "00000-000"),
        (19, "Tipo de declaração", "Original"),
        (20, "Residente no Brasil?", "Sim"),
    ]
    for row, label, value in labels_values:
        ws.cell(row, 2, label)
        ws.cell(row, 2).fill = fill(LIGHT)
        ws.cell(row, 2).font = Font(name="Aptos", size=9, bold=True, color=TEXT)
        ws.cell(row, 2).alignment = Alignment(vertical="center")
        ws.cell(row, 2).border = border()
        ws.cell(row, 3, value)
        style_input(ws.cell(row, 3), DATE_FMT if isinstance(value, date) else None)
        ws.row_dimensions[row].height = 25

    add_date_validation(ws, "C10", 1900, YEAR)
    add_list_validation(ws, "C14", "lista_estados_civis", "Selecione o estado civil.")
    add_list_validation(ws, "C17", "lista_ufs", "Selecione a unidade federativa.")
    add_list_validation(ws, "C19", "lista_tipos_declaracao", "Selecione Original ou Retificadora.")
    add_list_validation(ws, "C20", "lista_sim_nao", "Selecione Sim ou Não.")

    merge_box(ws, "F5:I5", "CONFERÊNCIA AUTOMÁTICA", bg=NAVY_2, color=WHITE, size=10, bold=True)
    checks = [
        (8, "Situação do cadastro", '=IF(COUNTA(C8:C20)>=10,"Cadastro consistente","Complete os campos")', None),
        (11, "Preenchimento", "=COUNTA(C8:C20)/13", PERCENT),
        (14, "Declaração", '=C19&" • Exercício "&Início!G6', None),
        (17, "Residência", '=IF(C20="Sim","Residente no Brasil","Revisar residência")', None),
    ]
    for row, label, formula, fmt in checks:
        merge_box(ws, f"F{row}:G{row+1}", label, bg=LIGHT, color=MUTED, size=9, bold=True)
        merge_box(ws, f"H{row}:I{row+1}", formula, bg=CALC, color=TEXT, size=10, bold=True,
                  number_format=fmt)

    merge_box(ws, "F21:I24",
              "PRIVACIDADE: nunca publique uma versão preenchida com seus dados reais. Faça uma cópia privada, mantenha backup protegido e deixe no GitHub somente o modelo demonstrativo.",
              bg=PALE_AMBER, color=TEXT, size=10, bold=True)
    merge_box(ws, "B23:C25", "VOLTAR AO PAINEL", bg=NAVY, color=WHITE, size=10, bold=True,
              hyperlink="#'Início'!A1")
    merge_box(ws, "B27:I29",
              "Os dados desta aba servem para organização e revisão. A transmissão da declaração deve ser feita somente pelos canais oficiais da Receita Federal.",
              bg=LIGHT, color=MUTED, size=9)
    ws.print_area = "A1:I29"


def setup_table_sheet(ws, title: str, subtitle: str, headers: list[str], widths: list[int],
                      table_name: str, samples: list[list], *, freeze: str = "A9") -> None:
    setup_sheet(ws, title, subtitle, freeze=freeze, zoom=80)
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + index)].width = width
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=len(headers))
    ws["A5"] = f"REGISTROS — {len(samples)} EXEMPLO(S) FICTÍCIO(S) PREENCHIDO(S)"
    style_section_header(ws["A5"], ws["A5"].value)
    ws["A5"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 25
    ws.merge_cells(start_row=6, start_column=1, end_row=6, end_column=len(headers))
    ws["A6"] = "Adicione novas linhas dentro da tabela. Células azul-claro são editáveis; colunas cinza contêm fórmulas."
    ws["A6"].fill = fill(LIGHT)
    ws["A6"].font = Font(name="Aptos", size=9, italic=True, color=MUTED)
    ws["A6"].alignment = Alignment(horizontal="center")

    for col, header in enumerate(headers, 1):
        cell = ws.cell(8, col, header)
        cell.fill = fill(NAVY)
        cell.font = Font(name="Aptos", size=8, bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border()
    ws.row_dimensions[8].height = 36

    for row in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        for col in range(1, len(headers) + 1):
            style_input(ws.cell(row, col))
        ws.row_dimensions[row].height = 30

    for row_offset, values in enumerate(samples, DATA_FIRST_ROW):
        for col, value in enumerate(values, 1):
            ws.cell(row_offset, col, value)

    end_col = chr(64 + len(headers))
    table = Table(displayName=table_name, ref=f"A8:{end_col}{DATA_LAST_ROW}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False,
    )
    ws.add_table(table)
    ws.auto_filter.ref = f"A8:{end_col}{DATA_LAST_ROW}"
    ws.print_title_rows = "1:8"
    ws.print_area = f"A1:{end_col}{DATA_LAST_ROW}"


def build_dependentes(wb: Workbook) -> None:
    ws = wb.create_sheet("Dependentes")
    headers = ["ID", "Nome", "CPF", "Nascimento", "Relação", "Reside com titular?", "Status", "Observações"]
    setup_table_sheet(
        ws, "DEPENDENTES", "Organize os dados cadastrais e marque os itens que ainda precisam de revisão.",
        headers, [13, 24, 19, 15, 24, 17, 14, 34], "tblDependentes", DEPENDENTES,
    )
    for row in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        ws.cell(row, 4).number_format = DATE_FMT
    add_date_validation(ws, f"D{DATA_FIRST_ROW}:D{DATA_LAST_ROW}", 1900, YEAR)
    add_list_validation(ws, f"E{DATA_FIRST_ROW}:E{DATA_LAST_ROW}", "lista_relacoes_dependente", "Selecione a relação.")
    add_list_validation(ws, f"F{DATA_FIRST_ROW}:F{DATA_LAST_ROW}", "lista_sim_nao", "Selecione Sim ou Não.")
    add_list_validation(ws, f"G{DATA_FIRST_ROW}:G{DATA_LAST_ROW}", "lista_status_cadastro", "Selecione o status cadastral.")
    ws.conditional_formatting.add(
        f"G{DATA_FIRST_ROW}:G{DATA_LAST_ROW}",
        FormulaRule(formula=[f'$G{DATA_FIRST_ROW}="Completo"'], fill=fill(PALE_GREEN)),
    )
    ws.conditional_formatting.add(
        f"G{DATA_FIRST_ROW}:G{DATA_LAST_ROW}",
        FormulaRule(formula=[f'$G{DATA_FIRST_ROW}="Revisar"'], fill=fill(PALE_AMBER)),
    )


def build_rendimentos(wb: Workbook) -> None:
    ws = wb.create_sheet("Rendimentos")
    headers = ["ID", "Categoria", "Fonte pagadora", "CPF/CNPJ", "Valor bruto", "IRRF", "INSS", "Documento", "Status", "Observações"]
    setup_table_sheet(
        ws, "RENDIMENTOS", "Centralize fontes pagadoras, valores e comprovantes sem calcular o imposto devido.",
        headers, [13, 24, 28, 20, 16, 15, 15, 28, 15, 30], "tblRendimentos", RENDIMENTOS,
    )
    for row in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        for col in (5, 6, 7):
            ws.cell(row, col).number_format = CURRENCY
    add_list_validation(ws, f"B{DATA_FIRST_ROW}:B{DATA_LAST_ROW}", "lista_categorias_rendimento", "Selecione a categoria do rendimento.")
    add_list_validation(ws, f"I{DATA_FIRST_ROW}:I{DATA_LAST_ROW}", "lista_status_documento", "Selecione o andamento da conferência.")
    for col in ("E", "F", "G"):
        add_decimal_validation(ws, f"{col}{DATA_FIRST_ROW}:{col}{DATA_LAST_ROW}")
    add_status_formatting(ws, "I")


def build_bens(wb: Workbook) -> None:
    ws = wb.create_sheet("Bens")
    headers = ["ID", "Grupo", "Código", "Descrição", "País", "Aquisição", "Situação anterior", "Situação atual", "Documento", "Status", "Observações"]
    setup_table_sheet(
        ws, "BENS E DIREITOS", "Compare os valores declarados no ano anterior com a situação atual e mantenha o documento de origem.",
        headers, [13, 25, 11, 30, 15, 15, 18, 18, 28, 15, 30], "tblBens", BENS,
    )
    for row in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        ws.cell(row, 6).number_format = DATE_FMT
        ws.cell(row, 7).number_format = CURRENCY
        ws.cell(row, 8).number_format = CURRENCY
    add_list_validation(ws, f"B{DATA_FIRST_ROW}:B{DATA_LAST_ROW}", "lista_grupos_bens", "Selecione o grupo do bem.")
    add_list_validation(ws, f"E{DATA_FIRST_ROW}:E{DATA_LAST_ROW}", "lista_paises", "Selecione o país.")
    add_list_validation(ws, f"J{DATA_FIRST_ROW}:J{DATA_LAST_ROW}", "lista_status_documento", "Selecione o andamento da conferência.")
    add_date_validation(ws, f"F{DATA_FIRST_ROW}:F{DATA_LAST_ROW}", 1900, YEAR)
    add_decimal_validation(ws, f"G{DATA_FIRST_ROW}:H{DATA_LAST_ROW}")
    add_status_formatting(ws, "J")


def build_pagamentos(wb: Workbook) -> None:
    ws = wb.create_sheet("Pagamentos")
    headers = ["ID", "Categoria", "Beneficiário", "CPF/CNPJ", "Pessoa", "Valor pago", "Reembolso", "Valor líquido", "Documento", "Status", "Observações"]
    setup_table_sheet(
        ws, "PAGAMENTOS", "Registre pagamentos e reembolsos; a coluna Valor líquido é calculada automaticamente.",
        headers, [13, 22, 27, 20, 15, 16, 16, 17, 28, 15, 30], "tblPagamentos", PAGAMENTOS,
    )
    for row in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        for col in (6, 7, 8):
            ws.cell(row, col).number_format = CURRENCY
        ws.cell(row, 8, f'=IF(COUNTA(F{row}:G{row})=0,"",MAX(0,F{row}-G{row}))')
        style_formula(ws.cell(row, 8), CURRENCY, "right")
    add_list_validation(ws, f"B{DATA_FIRST_ROW}:B{DATA_LAST_ROW}", "lista_categorias_pagamento", "Selecione a categoria do pagamento.")
    add_list_validation(ws, f"E{DATA_FIRST_ROW}:E{DATA_LAST_ROW}", "lista_titular_dependente", "Selecione Titular ou Dependente.")
    add_list_validation(ws, f"J{DATA_FIRST_ROW}:J{DATA_LAST_ROW}", "lista_status_documento", "Selecione o andamento da conferência.")
    add_decimal_validation(ws, f"F{DATA_FIRST_ROW}:G{DATA_LAST_ROW}")
    add_status_formatting(ws, "J")


def build_informes(wb: Workbook) -> None:
    ws = wb.create_sheet("Informes")
    headers = ["ID", "Categoria", "Fonte/Instituição", "Identificador", "Documento esperado", "Arquivo/Local", "Recebido em", "Status", "Observações"]
    setup_table_sheet(
        ws, "INFORMES", "Use esta tela como checklist central dos documentos necessários para a revisão.",
        headers, [13, 20, 28, 20, 27, 31, 16, 16, 32], "tblInformes", INFORMES,
    )
    for row in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        ws.cell(row, 7).number_format = DATE_FMT
    add_list_validation(ws, f"B{DATA_FIRST_ROW}:B{DATA_LAST_ROW}", "lista_categorias_informe", "Selecione a categoria do informe.")
    add_list_validation(ws, f"H{DATA_FIRST_ROW}:H{DATA_LAST_ROW}", "lista_status_documento", "Atualize o status do documento.")
    add_date_validation(ws, f"G{DATA_FIRST_ROW}:G{DATA_LAST_ROW}", YEAR, EXERCISE + 2)
    add_status_formatting(ws, "H")


def build_notas(wb: Workbook) -> None:
    ws = wb.create_sheet("Notas")
    headers = ["ID", "Categoria", "Assunto", "Nota", "Prioridade", "Prazo", "Status", "Alerta", "Responsável"]
    setup_table_sheet(
        ws, "NOTAS E PENDÊNCIAS", "Organize ações de conferência com prioridade, prazo e alerta automático.",
        headers, [13, 18, 29, 42, 14, 16, 17, 17, 18], "tblNotas", NOTAS,
    )
    for row in range(DATA_FIRST_ROW, DATA_LAST_ROW + 1):
        ws.cell(row, 6).number_format = DATE_FMT
        ws.cell(row, 8, f'=IF(A{row}="","",IF(G{row}="Concluída","OK",IF(AND(F{row}<>"",F{row}<TODAY()),"ATRASADA","ABERTA")))')
        style_formula(ws.cell(row, 8))
    add_list_validation(ws, f"B{DATA_FIRST_ROW}:B{DATA_LAST_ROW}", "lista_categorias_nota", "Selecione a categoria da nota.")
    add_list_validation(ws, f"E{DATA_FIRST_ROW}:E{DATA_LAST_ROW}", "lista_prioridades", "Selecione a prioridade.")
    add_list_validation(ws, f"G{DATA_FIRST_ROW}:G{DATA_LAST_ROW}", "lista_status_nota", "Selecione o status da ação.")
    add_date_validation(ws, f"F{DATA_FIRST_ROW}:F{DATA_LAST_ROW}", YEAR, EXERCISE + 5)
    ws.conditional_formatting.add(
        f"H{DATA_FIRST_ROW}:H{DATA_LAST_ROW}",
        FormulaRule(formula=[f'$H{DATA_FIRST_ROW}="ATRASADA"'], fill=fill(PALE_RED),
                    font=Font(color=RED, bold=True)),
    )
    ws.conditional_formatting.add(
        f"H{DATA_FIRST_ROW}:H{DATA_LAST_ROW}",
        FormulaRule(formula=[f'$H{DATA_FIRST_ROW}="OK"'], fill=fill(PALE_GREEN),
                    font=Font(color=EMERALD, bold=True)),
    )


def add_status_formatting(ws, column: str) -> None:
    target = f"{column}{DATA_FIRST_ROW}:{column}{DATA_LAST_ROW}"
    rules = [
        ("Conferido", PALE_GREEN, EMERALD),
        ("Recebido", INPUT, CYAN_DARK),
        ("Pendente", PALE_RED, RED),
        ("Não se aplica", CALC, MUTED),
    ]
    for text, bg, color in rules:
        ws.conditional_formatting.add(
            target,
            FormulaRule(formula=[f'${column}{DATA_FIRST_ROW}="{text}"'],
                        fill=fill(bg), font=Font(color=color, bold=True)),
        )


def build_guide(wb: Workbook) -> None:
    ws = wb.create_sheet("Guia")
    setup_sheet(ws, "GUIA DE USO", "Roteiro de preenchimento, privacidade e validação do IR Panorama.", freeze="A4", zoom=90)
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 82
    for col in "DEFGHIJKL":
        ws.column_dimensions[col].width = 3

    sections = [
        (5, "1. Antes de começar", "Baixe o modelo, salve uma cópia privada e mantenha o arquivo publicado apenas com dados fictícios. Defina o ano-calendário no painel."),
        (8, "2. Titular e dependentes", "Substitua os exemplos na cópia privada. Revise CPF, datas, relação de dependência e dados cadastrais com os documentos oficiais."),
        (11, "3. Rendimentos", "Inclua uma linha por fonte pagadora e transcreva valores apenas dos informes oficiais. A planilha soma valores; ela não calcula imposto devido."),
        (14, "4. Bens e direitos", "Use o grupo e o código apenas como organização preliminar. Compare os valores com a declaração anterior e confirme as regras oficiais antes de transmitir."),
        (17, "5. Pagamentos", "Informe valor pago e reembolso. Valor líquido = valor pago − reembolso, limitado a zero. Guarde recibos e identifique titular ou dependente."),
        (20, "6. Informes", "Atualize o status para Pendente, Recebido, Conferido ou Não se aplica. O painel calcula o percentual concluído e destaca pendências."),
        (23, "7. Notas", "Crie ações com prioridade e prazo. O alerta mostra ATRASADA quando a data passou e a ação ainda não foi concluída."),
        (26, "8. Links rápidos", "Os botões do painel abrem apenas páginas oficiais. Este projeto não usa API, não envia dados e não possui macro."),
        (29, "9. Conferência final", "Compare tudo com os documentos, a declaração anterior e o sistema oficial da Receita Federal. Procure orientação profissional em situações específicas."),
        (32, "10. Segurança", "Não armazene a cópia real em repositório público, pasta compartilhada ou serviço sem proteção. Use senha do dispositivo, backup criptografado e descarte seguro de cópias."),
    ]
    for row, title, body in sections:
        merge_box(ws, f"B{row}:B{row+1}", title, bg=NAVY_2, color=WHITE, size=9, bold=True,
                  align="left")
        merge_box(ws, f"C{row}:C{row+1}", body, bg=WHITE, color=TEXT, size=9, align="left")
        ws.row_dimensions[row].height = 28
        ws.row_dimensions[row + 1].height = 28

    merge_box(ws, "B35:C37",
              "IMPORTANTE: o IR Panorama é um organizador educacional e não substitui o programa oficial, a legislação vigente ou orientação de contador/advogado. Nenhum dado real foi incluído neste arquivo.",
              bg=PALE_AMBER, color=TEXT, size=10, bold=True)
    merge_box(ws, "B39:C40", f"Autoria: {AUTHOR}", bg=NAVY, color=CYAN, size=10, bold=True)
    merge_box(ws, "B42:C43", "VOLTAR AO PAINEL", bg=CYAN_DARK, color=WHITE, size=10, bold=True,
              hyperlink="#'Início'!A1")
    ws.print_area = "A1:C43"


def build_support(wb: Workbook) -> None:
    ws = wb.create_sheet("Apoio")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Listas internas para validação de dados — IR Panorama"
    ws["A1"].font = Font(name="Aptos Display", size=16, bold=True, color=WHITE)
    ws["A1"].fill = fill(NAVY)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(SUPPORT_LISTS))

    for col, (name, values) in enumerate(SUPPORT_LISTS.items(), 1):
        header = ws.cell(3, col, name)
        header.fill = fill(NAVY_2)
        header.font = Font(name="Aptos", size=8, bold=True, color=WHITE)
        header.alignment = Alignment(text_rotation=45, horizontal="center", vertical="bottom")
        ws.column_dimensions[chr(64 + col)].width = 23
        for row, value in enumerate(values, 4):
            ws.cell(row, col, value)
            ws.cell(row, col).fill = fill(LIGHT if row % 2 == 0 else WHITE)
            ws.cell(row, col).font = Font(name="Aptos", size=8, color=TEXT)
            ws.cell(row, col).border = border()
        end_row = 3 + len(values)
        col_letter = chr(64 + col)
        add_defined_name(wb, name, f"'Apoio'!${col_letter}$4:${col_letter}${end_row}")
    ws.sheet_state = "hidden"


def add_names(wb: Workbook) -> None:
    add_defined_name(wb, "ano_calendario", "'Início'!$C$6")
    add_defined_name(wb, "exercicio_declaracao", "'Início'!$G$6")
    add_defined_name(wb, "nome_titular", "'Titular'!$C$8")
    add_defined_name(wb, "cpf_titular", "'Titular'!$C$9")
    add_defined_name(wb, "dados_rendimentos", "'Rendimentos'!$A$9:$J$38")
    add_defined_name(wb, "dados_bens", "'Bens'!$A$9:$K$38")
    add_defined_name(wb, "dados_pagamentos", "'Pagamentos'!$A$9:$K$38")
    add_defined_name(wb, "dados_informes", "'Informes'!$A$9:$I$38")


def build_workbook() -> Workbook:
    wb = Workbook()
    wb.properties.title = "IR Panorama — Organizador de Declaração de Imposto de Renda"
    wb.properties.subject = "Projeto autoral para o desafio DIO de organização de dados fiscais no Excel"
    wb.properties.creator = AUTHOR
    wb.properties.lastModifiedBy = AUTHOR_HANDLE
    wb.properties.description = "Organizador local de titular, dependentes, rendimentos, bens, pagamentos, informes e pendências."
    wb.properties.keywords = "Excel, Imposto de Renda, organização, validação de dados, DIO"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"

    build_dashboard(wb)
    build_titular(wb)
    build_dependentes(wb)
    build_rendimentos(wb)
    build_bens(wb)
    build_pagamentos(wb)
    build_informes(wb)
    build_notas(wb)
    build_guide(wb)
    build_support(wb)
    add_names(wb)
    wb.active = 0
    return wb


def cached_values() -> dict[str, dict[str, float | int | str]]:
    total_income = sum(float(row[4]) for row in RENDIMENTOS)
    total_irrf = sum(float(row[5]) for row in RENDIMENTOS)
    total_inss = sum(float(row[6]) for row in RENDIMENTOS)
    total_assets = sum(float(row[7]) for row in BENS)
    payment_net = [max(0.0, float(row[5]) - float(row[6])) for row in PAGAMENTOS]
    total_payments = sum(payment_net)
    status_counts = {status: 0 for status in SUPPORT_LISTS["lista_status_documento"]}
    for row in INFORMES:
        status_counts[row[7]] += 1
    completed_docs = status_counts["Conferido"] + status_counts["Não se aplica"]
    progress = completed_docs / len(INFORMES)
    open_notes = sum(1 for row in NOTAS if row[6] != "Concluída")
    pending = (len(INFORMES) - completed_docs) + open_notes

    dashboard: dict[str, float | int | str] = {
        "G6": EXERCISE,
        "A10": total_income,
        "D10": total_assets,
        "G10": total_payments,
        "J10": progress,
        "A15": total_irrf,
        "D15": total_inss,
        "G15": pending,
        "J15": len(RENDIMENTOS) + len(BENS) + len(PAGAMENTOS),
        "A19": progress,
        "C35": "Pronto",
        "C37": f'{status_counts["Conferido"]} de {len(INFORMES)} conferidos',
        "C39": f"{open_notes} notas abertas",
        "C41": f'{status_counts["Pendente"]} documento(s) pendente(s)',
        "C43": "Revise as pendências",
        "O3": status_counts["Conferido"],
        "O4": status_counts["Recebido"],
        "O5": status_counts["Pendente"],
        "O6": status_counts["Não se aplica"],
    }
    titular = {
        "H8": "Cadastro consistente",
        "H11": 1.0,
        "H14": f"Original • Exercício {EXERCISE}",
        "H17": "Residente no Brasil",
    }
    pagamentos = {f"H{row}": value for row, value in enumerate(payment_net, DATA_FIRST_ROW)}
    notas = {
        "H9": "ATRASADA",
        "H10": "ATRASADA",
        "H11": "OK",
    }
    return {
        "xl/worksheets/sheet1.xml": dashboard,
        "xl/worksheets/sheet2.xml": titular,
        "xl/worksheets/sheet6.xml": pagamentos,
        "xl/worksheets/sheet8.xml": notas,
    }


def populate_formula_cache(path: Path) -> None:
    """Inclui resultados do exemplo no cache OOXML sem remover fórmulas."""
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", namespace)
    sheets = cached_values()
    temporary = path.with_name(path.stem + "-cache.xlsx")

    with ZipFile(path, "r") as source, ZipFile(temporary, "w", ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename in sheets:
                root = ET.fromstring(data)
                values = sheets[item.filename]
                for cell in root.findall(f".//{{{namespace}}}c"):
                    ref = cell.get("r")
                    if ref not in values:
                        continue
                    value = values[ref]
                    node = cell.find(f"{{{namespace}}}v")
                    if node is None:
                        node = ET.SubElement(cell, f"{{{namespace}}}v")
                    if isinstance(value, str):
                        cell.set("t", "str")
                        node.text = value
                    else:
                        if cell.get("t") == "str":
                            cell.attrib.pop("t", None)
                        node.text = format(value, ".15g")
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            target.writestr(item, data)
    temporary.replace(path)


def main() -> None:
    ROOT.mkdir(parents=True, exist_ok=True)
    wb = build_workbook()
    wb.save(OUTPUT)
    populate_formula_cache(OUTPUT)
    print(f"Planilha criada: {OUTPUT}")


if __name__ == "__main__":
    main()
