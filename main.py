import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import matplotlib.pyplot as plt
from openpyxl import Workbook

def calcular_parcelas():
    try:
        valor_total = float(entrada_valor_total.get())
        meses = int(entrada_meses.get())
        taxa_anual = float(entrada_juros.get())

        # ✅ Usar taxa mensal equivalente (composta)
        taxa_mensal = (1 + taxa_anual / 100) ** (1 / 12) - 1

        if entrada_valor_entrada.get():
            valor_entrada = float(entrada_valor_entrada.get())
        elif entrada_percentual_entrada.get():
            percentual = float(entrada_percentual_entrada.get())
            valor_entrada = valor_total * percentual / 100
        else:
            messagebox.showerror("Erro", "Informe o valor de entrada ou a porcentagem.")
            return

        valor_financiado = valor_total - valor_entrada
        sistema = sistema_var.get()

        parcelas = []
        saldo_devedor = valor_financiado

        if sistema == "SAC":
            amortizacao = valor_financiado / meses
            for i in range(1, meses + 1):
                juros = saldo_devedor * taxa_mensal
                parcela = amortizacao + juros
                parcelas.append((i, round(parcela, 2)))
                saldo_devedor -= amortizacao

        elif sistema == "PRICE":
            parcela_fixa = valor_financiado * (taxa_mensal * (1 + taxa_mensal) ** meses) / ((1 + taxa_mensal) ** meses - 1)
            saldo = valor_financiado
            for i in range(1, meses + 1):
                juros = saldo * taxa_mensal
                amort = parcela_fixa - juros
                saldo -= amort
                parcelas.append((i, round(parcela_fixa, 2)))

        # ✅ Calcular total pago
        total_pago = sum([valor for _, valor in parcelas])

        texto_resultado = f"Valor financiado: R$ {valor_financiado:,.2f}\nSistema: {sistema}\n"
        texto_resultado += f"Taxa mensal equivalente usada: {taxa_mensal * 100:.4f}%\n\n"
        texto_resultado += "Parcela     Valor\n" + "-"*25 + "\n"
        for num, valor in parcelas:
            texto_resultado += f"{num:>3}        R$ {valor:,.2f}\n"

        texto_resultado += f"\nTotal pago ao final do financiamento: R$ {total_pago:,.2f}\n"

        resultado_texto.delete("1.0", tk.END)
        resultado_texto.insert(tk.END, texto_resultado)

        global dados_parcelas
        dados_parcelas = parcelas

    except ValueError:
        messagebox.showerror("Erro", "Preencha todos os campos corretamente.")


def exportar_excel():
    if not dados_parcelas:
        messagebox.showwarning("Aviso", "Calcule as parcelas antes de exportar.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Planilhas Excel", "*.xlsx")])
    if not file_path:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Parcelas"
    ws.append(["Número da Parcela", "Valor (R$)"])
    for num, valor in dados_parcelas:
        ws.append([num, valor])
    wb.save(file_path)
    messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{file_path}")

def mostrar_grafico():
    if not dados_parcelas:
        messagebox.showwarning("Aviso", "Calcule as parcelas antes de visualizar o gráfico.")
        return

    x = [p[0] for p in dados_parcelas]
    y = [p[1] for p in dados_parcelas]

    plt.figure(figsize=(10, 5))
    plt.plot(x, y, marker='o')
    plt.title(f"Evolução das Parcelas - {sistema_var.get()}")
    plt.xlabel("Número da Parcela")
    plt.ylabel("Valor da Parcela (R$)")
    plt.grid(True)
    plt.tight_layout()
    plt.show()

# Variável global
dados_parcelas = []

# Interface
root = tk.Tk()
root.title("Simulador de Financiamento")

frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0)

ttk.Label(frame, text="Valor total do imóvel (R$):").grid(row=0, column=0, sticky="w")
entrada_valor_total = ttk.Entry(frame, width=20)
entrada_valor_total.grid(row=0, column=1)

ttk.Label(frame, text="Valor de entrada (R$):").grid(row=1, column=0, sticky="w")
entrada_valor_entrada = ttk.Entry(frame, width=20)
entrada_valor_entrada.grid(row=1, column=1)

ttk.Label(frame, text="Ou % de entrada (%):").grid(row=2, column=0, sticky="w")
entrada_percentual_entrada = ttk.Entry(frame, width=20)
entrada_percentual_entrada.grid(row=2, column=1)

ttk.Label(frame, text="Número de meses:").grid(row=3, column=0, sticky="w")
entrada_meses = ttk.Entry(frame, width=20)
entrada_meses.grid(row=3, column=1)

ttk.Label(frame, text="Taxa de juros anual (%):").grid(row=4, column=0, sticky="w")
entrada_juros = ttk.Entry(frame, width=20)
entrada_juros.grid(row=4, column=1)

ttk.Label(frame, text="Sistema:").grid(row=5, column=0, sticky="w")
sistema_var = tk.StringVar(value="SAC")
ttk.Radiobutton(frame, text="SAC", variable=sistema_var, value="SAC").grid(row=5, column=1, sticky="w")
ttk.Radiobutton(frame, text="PRICE", variable=sistema_var, value="PRICE").grid(row=5, column=1, sticky="e")

ttk.Button(frame, text="Calcular", command=calcular_parcelas).grid(row=6, column=0, pady=10)
ttk.Button(frame, text="Exportar para Excel", command=exportar_excel).grid(row=6, column=1, pady=10)
ttk.Button(frame, text="Mostrar Gráfico", command=mostrar_grafico).grid(row=7, column=0, columnspan=2)

resultado_texto = tk.Text(root, height=20, width=60)
resultado_texto.grid(row=1, column=0, padx=10, pady=10)

root.mainloop()
