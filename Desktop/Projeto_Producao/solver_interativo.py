# ===========================================================
# SOFTWARE DE PESQUISA OPERACIONAL - SOLVER PYTHON
# VERSÃO INTERATIVA COM MENU
# ===========================================================

from pulp import LpProblem, LpVariable, LpMaximize, LpMinimize, lpSum, PULP_CBC_CMD
import pandas as pd
import matplotlib.pyplot as plt
import os

# -----------------------------------------------------------
# Função para criar e resolver o modelo
# -----------------------------------------------------------
def resolver_problema():
    print("\n==========================================================")
    print("         CRIAÇÃO DE UM NOVO PROBLEMA")
    print("==========================================================")

    # Tipo de problema
    print("\n1️⃣ Escolha o tipo de problema:")
    print("   [1] Maximização")
    print("   [2] Minimização")
    tipo = input("Digite 1 ou 2: ")

    if tipo == "1":
        problema = LpProblem("Problema_PO", LpMaximize)
    else:
        problema = LpProblem("Problema_PO", LpMinimize)

    # Variáveis de decisão
    n_var = int(input("\nQuantas variáveis de decisão existem? "))
    coef_objetivo = []
    variaveis = []

    for i in range(n_var):
        nome = f"x{i+1}"
        var = LpVariable(nome, lowBound=0)  # não-negatividade
        variaveis.append(var)
        c = float(input(f"Coeficiente da variável {nome} na função objetivo: "))
        coef_objetivo.append(c)

    # Função objetivo
    problema += lpSum([coef_objetivo[i] * variaveis[i] for i in range(n_var)]), "Função_Objetivo"

    # Restrições
    n_rest = int(input("\nQuantas restrições o problema possui? "))

    for r in range(n_rest):
        print(f"\n--- Restrição {r+1} ---")
        coefs = []
        for i in range(n_var):
            c = float(input(f"Coeficiente de x{i+1}: "))
            coefs.append(c)
        sinal = input("Sinal da restrição (<= , = , >=): ")
        rhs = float(input("Valor do lado direito da restrição: "))

        if sinal == "<=":
            problema += lpSum([coefs[i] * variaveis[i] for i in range(n_var)]) <= rhs
        elif sinal == ">=":
            problema += lpSum([coefs[i] * variaveis[i] for i in range(n_var)]) >= rhs
        else:
            problema += lpSum([coefs[i] * variaveis[i] for i in range(n_var)]) == rhs

    print("\nModelo criado com sucesso! Resolvendo...\n")

    # Resolver
    problema.solve(PULP_CBC_CMD(msg=0))

    print("Status:", problema.status)
    print("Valor ótimo (função objetivo):", problema.objective.value())

    print("\nValores das variáveis:")
    for v in problema.variables():
        print(f"{v.name} = {v.value()}")

    # Criar DataFrame de resultado
    dados = {
        "Variável": [v.name for v in problema.variables()],
        "Valor ótimo": [v.value() for v in problema.variables()]
    }
    df = pd.DataFrame(dados)
    df.loc[len(df)] = ["Função Objetivo", problema.objective.value()]
    df.to_excel("relatorio_solver.xlsx", index=False)

    # Gráfico
    plt.figure()
    plt.bar(df["Variável"], df["Valor ótimo"], color="teal")
    plt.title("Resultados do Modelo - Solver Python")
    plt.xlabel("Variáveis")
    plt.ylabel("Valor ótimo")
    plt.tight_layout()
    plt.show()

    print("\n✅ Relatório salvo como 'relatorio_solver.xlsx'")
    print("==========================================================\n")


# -----------------------------------------------------------
# Função para visualizar relatório existente
# -----------------------------------------------------------
def ver_relatorio():
    if not os.path.exists("relatorio_solver.xlsx"):
        print("\n⚠️ Nenhum relatório encontrado! Resolva um problema primeiro.\n")
        return
    df = pd.read_excel("relatorio_solver.xlsx")
    print("\nÚltimo relatório salvo:\n")
    print(df)
    print("\nValor ótimo:", df.loc[df['Variável'] == 'Função Objetivo', 'Valor ótimo'].values[0])


# -----------------------------------------------------------
# Menu principal
# -----------------------------------------------------------
def menu():
    while True:
        print("\n==========================================================")
        print("        SOFTWARE DE PESQUISA OPERACIONAL - SOLVER")
        print("==========================================================")
        print("[1] Criar e resolver novo problema")
        print("[2] Ver último relatório salvo")
        print("[3] Sair")
        print("==========================================================")

        opcao = input("Escolha uma opção: ")

        if opcao == "1":
            resolver_problema()
        elif opcao == "2":
            ver_relatorio()
        elif opcao == "3":
            print("\nEncerrando o programa... 👋\n")
            break
        else:
            print("\nOpção inválida! Tente novamente.\n")


# -----------------------------------------------------------
# Execução do programa
# -----------------------------------------------------------
if __name__ == "__main__":
    menu()
