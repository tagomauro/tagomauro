import tkinter as tk
from tkinter import messagebox
import openpyxl
from pathlib import Path

ARQUIVO = Path(__file__).with_name("Seções.xlsx")


def carregar_secoes():
    wb = openpyxl.load_workbook(ARQUIVO, data_only=True)
    ws = wb.active

    agregacoes = {}
    secoes_verdes = set()

    for row in ws.iter_rows():
        linha = row[0].row
        numero_verde = ws.cell(row=linha, column=5).value

        if numero_verde is not None:
            numero_verde = str(numero_verde).strip()
            secoes_verdes.add(numero_verde)

            for cell in row[:4]:
                cor = cell.fill.fgColor

                if (
                    cell.value is not None
                    and cell.fill.fill_type == "solid"
                    and cor.type == "rgb"
                    and cor.rgb == "FFFFFF00"
                ):
                    agregacoes[str(cell.value).strip()] = numero_verde

    return agregacoes, secoes_verdes


try:
    AGREGACOES, SECOES_VERDES = carregar_secoes()
except Exception as e:
    AGREGACOES, SECOES_VERDES = {}, set()
    ERRO_PLANILHA = str(e)
else:
    ERRO_PLANILHA = None


def mostrar_consulta():
    numero = entrada.get().strip()

    if not numero:
        resultado.config(
            text="Digite o número de sua seção.",
            fg="#9C0006"
        )
        return

    if numero in AGREGACOES:
        resultado.config(
            text=f"Sua seção foi agregada pela seção {AGREGACOES[numero]}.",
            fg="#9C0006"
        )
    elif numero in SECOES_VERDES:
        resultado.config(
            text="Sua seção não foi agregada.",
            fg="#006100"
        )
    else:
        resultado.config(
            text="Número de seção não encontrado na planilha.",
            fg="#9C6500"
        )


def sim():
    pergunta.config(text="Digite o número de sua seção:")
    entrada.pack(pady=(8, 12))
    botao_consultar.pack(pady=(0, 15))
    entrada.focus_set()
    botao_sim.pack_forget()
    botao_nao.pack_forget()


def nao():
    resultado.config(
        text="Para continuar, é necessário informar o número de sua seção.",
        fg="#444444"
    )


janela = tk.Tk()
janela.title("Consulta de Seção Eleitoral")
janela.geometry("620x420")
janela.resizable(False, False)
janela.configure(bg="#F4F6F8")

titulo = tk.Label(
    janela,
    text="CONSULTA DE SEÇÃO ELEITORAL",
    font=("Arial", 20, "bold"),
    bg="#F4F6F8",
    fg="#1F2937"
)
titulo.pack(pady=(35, 20))

pergunta = tk.Label(
    janela,
    text="Sabe o número de sua seção?",
    font=("Arial", 15),
    bg="#F4F6F8",
    fg="#1F2937"
)
pergunta.pack()

frame_botoes = tk.Frame(janela, bg="#F4F6F8")
frame_botoes.pack(pady=20)

botao_sim = tk.Button(
    frame_botoes,
    text="SIM",
    command=sim,
    width=12,
    height=2,
    font=("Arial", 12, "bold"),
    cursor="hand2"
)
botao_sim.pack(side="left", padx=10)

botao_nao = tk.Button(
    frame_botoes,
    text="NÃO",
    command=nao,
    width=12,
    height=2,
    font=("Arial", 12, "bold"),
    cursor="hand2"
)
botao_nao.pack(side="left", padx=10)

entrada = tk.Entry(
    janela,
    font=("Arial", 18),
    justify="center",
    width=12
)

botao_consultar = tk.Button(
    janela,
    text="CONSULTAR",
    command=mostrar_consulta,
    width=16,
    height=2,
    font=("Arial", 11, "bold"),
    cursor="hand2"
)

resultado = tk.Label(
    janela,
    text="",
    font=("Arial", 16, "bold"),
    bg="#F4F6F8",
    wraplength=550,
    justify="center"
)
resultado.pack(pady=20)

if ERRO_PLANILHA:
    resultado.config(
        text="Não foi possível carregar a planilha Seções.xlsx.",
        fg="#9C0006"
    )
    botao_sim.config(state="disabled")
    botao_nao.config(state="disabled")

janela.mainloop()
