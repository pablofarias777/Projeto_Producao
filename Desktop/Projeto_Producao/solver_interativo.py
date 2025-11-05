import tkinter as tk
from tkinter import ttk, messagebox
from pulp import LpProblem, LpVariable, LpMaximize, LpMinimize, lpSum, PULP_CBC_CMD, LpStatus
import pandas as pd
import os


class SolverGUI:
    def __init__(self, master):
        self.master = master
        master.title("Solver de Pesquisa Operacional")

        # Variáveis de controle
        self.tipo_problema = tk.StringVar(value="max")
        self.n_var = tk.IntVar(value=4)   # número de produtos (x1, x2, ...)
        self.n_rest = tk.IntVar(value=4)  # número de restrições de capacidade

        # Referências para widgets dinâmicos
        self.obj_coef_entries = []     # coeficientes da função objetivo (x)
        self.folga_coef_entries = []   # coeficientes das variáveis s (penalidades)
        self.demanda_entries = []      # demandas de cada produto
        self.rest_coef_entries = []    # matriz [restrição][xj]
        self.rest_sinal_vars = []      # operadores <=, =, >=
        self.rest_rhs_entries = []     # lado direito das restrições de capacidade

        # ---------- FRAME CONFIGURAÇÃO INICIAL ----------
        frame_config = ttk.LabelFrame(master, text="1. Definir problema")
        frame_config.pack(fill="x", padx=10, pady=10)

        # Tipo de problema
        ttk.Label(frame_config, text="Função objetivo:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Radiobutton(
            frame_config, text="Maximizar", variable=self.tipo_problema, value="max"
        ).grid(row=0, column=1, sticky="w", padx=5, pady=5)
        ttk.Radiobutton(
            frame_config, text="Minimizar", variable=self.tipo_problema, value="min"
        ).grid(row=0, column=2, sticky="w", padx=5, pady=5)

        # Número de variáveis (produtos)
        ttk.Label(frame_config, text="Nº de variáveis de decisão (x):").grid(
            row=1, column=0, sticky="w", padx=5, pady=5
        )
        ttk.Entry(frame_config, textvariable=self.n_var, width=5).grid(
            row=1, column=1, sticky="w", padx=5, pady=5
        )

        # Número de restrições de capacidade
        ttk.Label(frame_config, text="Nº de restrições de capacidade:").grid(
            row=2, column=0, sticky="w", padx=5, pady=5
        )
        ttk.Entry(frame_config, textvariable=self.n_rest, width=5).grid(
            row=2, column=1, sticky="w", padx=5, pady=5
        )

        ttk.Button(
            frame_config, text="Criar modelo", command=self.criar_campos_modelo
        ).grid(row=3, column=0, columnspan=3, pady=10)

        # ---------- FRAME CAMPOS DO MODELO ----------
        self.frame_modelo = ttk.LabelFrame(master, text="2. Modelo")
        self.frame_modelo.pack(fill="both", expand=True, padx=10, pady=5)

        # ---------- FRAME RESULTADO ----------
        frame_result = ttk.LabelFrame(master, text="3. Resultado")
        frame_result.pack(fill="both", expand=True, padx=10, pady=10)

        self.text_resultado = tk.Text(frame_result, height=12)
        self.text_resultado.pack(fill="both", expand=True, padx=5, pady=5)

    # ---------------------------------------------------
    # Cria os campos de função objetivo, demandas e restrições
    # ---------------------------------------------------
    def criar_campos_modelo(self):
        # Apaga conteúdo anterior
        for widget in self.frame_modelo.winfo_children():
            widget.destroy()

        self.obj_coef_entries = []
        self.folga_coef_entries = []
        self.demanda_entries = []
        self.rest_coef_entries = []
        self.rest_sinal_vars = []
        self.rest_rhs_entries = []

        n_var = self.n_var.get()
        n_rest = self.n_rest.get()

        if n_var <= 0 or n_rest <= 0:
            messagebox.showerror("Erro", "Informe números positivos de variáveis e restrições.")
            return

        # ---------- Função Objetivo (coeficientes de x) ----------
        frame_obj = ttk.LabelFrame(self.frame_modelo, text="Função Objetivo (Z)")
        frame_obj.pack(fill="x", padx=5, pady=5)

        ttk.Label(frame_obj, text="Z = ").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        for j in range(n_var):
            ttk.Label(frame_obj, text=f"coef x{j+1}:").grid(row=0, column=2*j+1, padx=2, pady=5)
            entry = ttk.Entry(frame_obj, width=6)
            entry.grid(row=0, column=2*j+2, padx=2, pady=5)
            self.obj_coef_entries.append(entry)

        # ---------- Coeficientes das variáveis de excesso/falta (s) ----------
        frame_folga = ttk.LabelFrame(
            self.frame_modelo,
            text="Coef. das variáveis de exceção (s1, s2, ...)"
        )
        frame_folga.pack(fill="x", padx=5, pady=5)

        for j in range(n_var):
            ttk.Label(frame_folga, text=f"coef s{j+1}:").grid(row=0, column=2*j+1, padx=2, pady=5)
            entry = ttk.Entry(frame_folga, width=6)
            entry.grid(row=0, column=2*j+2, padx=2, pady=5)
            self.folga_coef_entries.append(entry)

        # ---------- Demandas dos produtos (equações xj + sj = demanda) ----------
        frame_dem = ttk.LabelFrame(self.frame_modelo, text="Demandas dos produtos (xj + sj = demanda)")
        frame_dem.pack(fill="x", padx=5, pady=5)

        for j in range(n_var):
            ttk.Label(frame_dem, text=f"demanda de x{j+1}:").grid(row=0, column=2*j, padx=2, pady=5)
            entry = ttk.Entry(frame_dem, width=6)
            entry.grid(row=0, column=2*j+1, padx=2, pady=5)
            self.demanda_entries.append(entry)

        # ---------- Restrições de capacidade ----------
        frame_rest = ttk.LabelFrame(self.frame_modelo, text="Restrições de capacidade")
        frame_rest.pack(fill="x", padx=5, pady=5)

        # Cabeçalho
        for j in range(n_var):
            ttk.Label(frame_rest, text=f"x{j+1}").grid(row=0, column=j, padx=2, pady=2)
        ttk.Label(frame_rest, text="Sinal").grid(row=0, column=n_var, padx=5, pady=2)
        ttk.Label(frame_rest, text="LD").grid(row=0, column=n_var + 1, padx=5, pady=2)

        for r in range(n_rest):
            linha_entries = []
            for j in range(n_var):
                e = ttk.Entry(frame_rest, width=6)
                e.grid(row=r+1, column=j, padx=2, pady=2)
                linha_entries.append(e)
            self.rest_coef_entries.append(linha_entries)

            sinal_var = tk.StringVar(value="<=")
            cb = ttk.Combobox(frame_rest, textvariable=sinal_var, width=4, state="readonly")
            cb["values"] = ("<=", "=", ">=")
            cb.grid(row=r+1, column=n_var, padx=2, pady=2)
            self.rest_sinal_vars.append(sinal_var)

            rhs_entry = ttk.Entry(frame_rest, width=8)
            rhs_entry.grid(row=r+1, column=n_var + 1, padx=2, pady=2)
            self.rest_rhs_entries.append(rhs_entry)

        # Botão Resolver
        ttk.Button(
            self.frame_modelo, text="Resolver modelo", command=self.resolver_modelo
        ).pack(pady=10)

    # ---------------------------------------------------
    # Monta o modelo no PuLP e resolve
    # ---------------------------------------------------
    def resolver_modelo(self):
        try:
            n_var = self.n_var.get()
            n_rest = self.n_rest.get()

            # Coeficientes da função objetivo (x)
            coef_obj = [float(self.obj_coef_entries[j].get()) for j in range(n_var)]

            # Coeficientes das variáveis s (penalidades)
            coef_s = [float(self.folga_coef_entries[j].get()) for j in range(n_var)]

            # Demandas
            demandas = [float(self.demanda_entries[j].get()) for j in range(n_var)]

            # Cria modelo
            if self.tipo_problema.get() == "max":
                problema = LpProblem("Problema_PO", LpMaximize)
            else:
                problema = LpProblem("Problema_PO", LpMinimize)

            # Variáveis de decisão (x) e de excesso/falta (s)
            x = [LpVariable(f"x{j+1}", lowBound=0) for j in range(n_var)]
            s = [LpVariable(f"s{j+1}", lowBound=0) for j in range(n_var)]

            # Função objetivo: sum(cj*xj) - sum(pj*sj)
            problema += (
                lpSum(coef_obj[j] * x[j] for j in range(n_var))
                - lpSum(coef_s[j] * s[j] for j in range(n_var))
            ), "Funcao_Objetivo"

            # Restrições de capacidade (somente x)
            for r in range(n_rest):
                coefs_r = []
                for j in range(n_var):
                    txt = self.rest_coef_entries[r][j].get().strip()
                    valor = float(txt) if txt != "" else 0.0
                    coefs_r.append(valor)

                sinal = self.rest_sinal_vars[r].get()
                rhs = float(self.rest_rhs_entries[r].get())

                expr = lpSum(coefs_r[j] * x[j] for j in range(n_var))

                if sinal == "<=":
                    problema += expr <= rhs
                elif sinal == ">=":
                    problema += expr >= rhs
                else:
                    problema += expr == rhs

            # Restrições de demanda: xj + sj = demanda_j
            for j in range(n_var):
                problema += x[j] + s[j] == demandas[j], f"demanda_{j+1}"

            # Resolve
            problema.solve(PULP_CBC_CMD(msg=0))

            status = LpStatus[problema.status]
            valor_obj = problema.objective.value()

            # Monta saída
            saida = []
            saida.append(f"Status: {status}")
            saida.append(f"Z = {valor_obj:.3f}\n")

            linhas_df = []

            for v in problema.variables():
                saida.append(f"{v.name} = {v.value()}")
                linhas_df.append({"Variável": v.name, "Valor ótimo": v.value()})

            linhas_df.append({"Variável": "Z (Função Objetivo)", "Valor ótimo": valor_obj})

            # Atualiza painel de texto
            self.text_resultado.delete("1.0", tk.END)
            self.text_resultado.insert(tk.END, "\n".join(saida))

            # Salva em Excel
            df = pd.DataFrame(linhas_df)
            relatorio_nome = "relatorio_solver.xlsx"
            df.to_excel(relatorio_nome, index=False)

            messagebox.showinfo(
                "Concluído",
                f"Modelo resolvido com sucesso!\nRelatório salvo em {os.path.abspath(relatorio_nome)}"
            )

        except ValueError:
            messagebox.showerror(
                "Erro de entrada",
                "Verifique se todos os campos numéricos foram preenchidos corretamente."
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao resolver o modelo:\n{e}")


# -----------------------------------------------------------
# Execução
# -----------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = SolverGUI(root)
    root.mainloop()
