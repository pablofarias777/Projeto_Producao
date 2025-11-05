import tkinter as tk
from tkinter import ttk, messagebox
from pulp import *
import pandas as pd

class SolverGUI:
    def __init__(self, master):
        self.master = master
        master.title("Solver de Pesquisa Operacional")

        self.tipo_problema = tk.StringVar(value="max")
        self.n_var = tk.IntVar(value=4)
        self.n_rest = tk.IntVar(value=4)

        self.obj_coef_entries = []
        self.rest_coef_entries = []
        self.rest_sinal_vars = []
        self.rest_rhs_entries = []
        self.folga_coef_entries = []

        frame_config = ttk.LabelFrame(master, text="1. Configuração")
        frame_config.pack(fill="x", padx=10, pady=10)

        ttk.Radiobutton(frame_config, text="Maximizar", variable=self.tipo_problema, value="max").grid(row=0, column=1)
        ttk.Radiobutton(frame_config, text="Minimizar", variable=self.tipo_problema, value="min").grid(row=0, column=2)

        ttk.Label(frame_config, text="Nº de variáveis:").grid(row=1, column=0)
        ttk.Entry(frame_config, textvariable=self.n_var, width=5).grid(row=1, column=1)

        ttk.Label(frame_config, text="Nº de restrições:").grid(row=2, column=0)
        ttk.Entry(frame_config, textvariable=self.n_rest, width=5).grid(row=2, column=1)

        ttk.Button(frame_config, text="Criar modelo", command=self.criar_campos_modelo).grid(row=3, column=0, columnspan=3, pady=5)

        self.frame_modelo = ttk.LabelFrame(master, text="2. Modelo")
        self.frame_modelo.pack(fill="both", expand=True, padx=10, pady=5)

        frame_result = ttk.LabelFrame(master, text="3. Resultado")
        frame_result.pack(fill="both", expand=True, padx=10, pady=10)
        self.text_resultado = tk.Text(frame_result, height=10)
        self.text_resultado.pack(fill="both", expand=True, padx=5, pady=5)

    def criar_campos_modelo(self):
        for w in self.frame_modelo.winfo_children():
            w.destroy()

        self.obj_coef_entries = []
        self.rest_coef_entries = []
        self.rest_sinal_vars = []
        self.rest_rhs_entries = []
        self.folga_coef_entries = []

        n_var = self.n_var.get()
        n_rest = self.n_rest.get()

        frame_obj = ttk.LabelFrame(self.frame_modelo, text="Função Objetivo (Z)")
        frame_obj.pack(fill="x", padx=5, pady=5)
        ttk.Label(frame_obj, text="Z = ").grid(row=0, column=0)
        for j in range(n_var):
            ttk.Label(frame_obj, text=f"x{j+1}:").grid(row=0, column=2*j+1)
            e = ttk.Entry(frame_obj, width=6)
            e.grid(row=0, column=2*j+2)
            self.obj_coef_entries.append(e)

        frame_folga = ttk.LabelFrame(self.frame_modelo, text="Coef. das variáveis de exceção (s1, s2, ...)")
        frame_folga.pack(fill="x", padx=5, pady=5)
        for j in range(n_rest):
            ttk.Label(frame_folga, text=f"s{j+1}:").grid(row=0, column=2*j)
            e = ttk.Entry(frame_folga, width=6)
            e.grid(row=0, column=2*j+1)
            self.folga_coef_entries.append(e)

        frame_rest = ttk.LabelFrame(self.frame_modelo, text="Restrições")
        frame_rest.pack(fill="x", padx=5, pady=5)

        for j in range(n_var):
            ttk.Label(frame_rest, text=f"x{j+1}").grid(row=0, column=j)
        ttk.Label(frame_rest, text="Sinal").grid(row=0, column=n_var)
        ttk.Label(frame_rest, text="LD").grid(row=0, column=n_var+1)

        for r in range(n_rest):
            linha = []
            for j in range(n_var):
                e = ttk.Entry(frame_rest, width=6)
                e.grid(row=r+1, column=j)
                linha.append(e)
            self.rest_coef_entries.append(linha)

            s = tk.StringVar(value="<=")
            cb = ttk.Combobox(frame_rest, textvariable=s, values=["<=", ">=", "="], width=4, state="readonly")
            cb.grid(row=r+1, column=n_var)
            self.rest_sinal_vars.append(s)

            rhs = ttk.Entry(frame_rest, width=8)
            rhs.grid(row=r+1, column=n_var+1)
            self.rest_rhs_entries.append(rhs)

        ttk.Button(self.frame_modelo, text="Resolver", command=self.resolver_modelo).pack(pady=10)

    def resolver_modelo(self):
        try:
            n_var = self.n_var.get()
            n_rest = self.n_rest.get()

            coef_x = [float(e.get()) for e in self.obj_coef_entries]
            coef_s = [float(e.get()) for e in self.folga_coef_entries]

            modelo = LpProblem("Modelo", LpMaximize if self.tipo_problema.get()=="max" else LpMinimize)

            x = [LpVariable(f"x{j+1}", lowBound=0) for j in range(n_var)]
            s = [LpVariable(f"s{j+1}", lowBound=0) for j in range(n_rest)]

            # ✅ Corrigido: subtrai os coeficientes negativos das variáveis de exceção
            modelo += lpSum(coef_x[j]*x[j] for j in range(n_var)) + lpSum(coef_s[j]*s[j] for j in range(n_rest)), "Z"

            # ✅ Restrições corrigidas
            for i in range(n_rest):
                coefs = [float(e.get() or 0) for e in self.rest_coef_entries[i]]
                rhs = float(self.rest_rhs_entries[i].get() or 0)
                sinal = self.rest_sinal_vars[i].get()
                expr = lpSum(coefs[j]*x[j] for j in range(n_var))

                if sinal == "<=":
                    modelo += expr + s[i] == rhs
                elif sinal == ">=":
                    modelo += expr - s[i] == rhs
                else:
                    modelo += expr == rhs

            modelo.solve(PULP_CBC_CMD(msg=0))

            z_val = value(modelo.objective)

            saida = [f"Status: {LpStatus[modelo.status]}", f"Z = {z_val:.3f}\n"]
            dados = []
            for v in modelo.variables():
                saida.append(f"{v.name} = {v.varValue}")
                dados.append({"Variável": v.name, "Valor": v.varValue})
            dados.append({"Variável": "Z", "Valor": z_val})

            df = pd.DataFrame(dados)
            df.to_excel("resultado_solver.xlsx", index=False)

            self.text_resultado.delete("1.0", tk.END)
            self.text_resultado.insert(tk.END, "\n".join(saida))

            messagebox.showinfo("Sucesso", f"Z = {z_val:.3f}\nResultado salvo em resultado_solver.xlsx")

        except Exception as e:
            messagebox.showerror("Erro", f"{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SolverGUI(root)
    root.mainloop()