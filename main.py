import os
import argparse
from tenable.io import TenableIO
from pystyle import *
import random

from config import CLIENTES
from exporters import exportar_assets, exportar_agents
from utils import validar_credenciais, padronizar_excel_agents, banners, banner,banner1


def main():
    parser = argparse.ArgumentParser(description="Exportador de Assets e Agents da Tenable")
    parser.add_argument('-c', '--cliente', help='Nome do cliente', required=False)
    parser.add_argument('-t', '--tipo', choices=['assets', 'agents'], help='Tipo de exportação', required=False)
    parser.add_argument('-f', '--filtro', choices=['offline', 'nogroup', 'all', 'compare'], help='Filtro para agentes', required=False)
    parser.add_argument('-p', '--padronizar',help='Caminho do arquivo CSV ou XLSX para padronizar a planilha Agents',required=False, metavar='ARQUIVO')

    args = parser.parse_args()

    selected_banner = random.choice(banners)
    if selected_banner == banner:
        print(Colorate.DiagonalBackwards(Colors.blue_to_cyan, "{}".format(selected_banner), 2))
    elif selected_banner == banner1:
        print(Colorate.DiagonalBackwards(Colors.red_to_black, "{}".format(selected_banner), 2))
    else:
        print(selected_banner)
    print("[...]          Tenable Assets & Agents Exporter       [...]\n")

    if args.padronizar:
        padronizar_excel_agents(args.padronizar)
        return

    if not args.cliente:
        print("Clientes disponíveis:")
        for i, cliente in enumerate(CLIENTES):
            print(f"  [{i + 1}] {cliente}")
        print("  [0] Sair")
        print("\n  [99] Padronizar Excel (Agents)")

        try:
            escolha = int(input("\nEscolha um cliente: "))
            if escolha == 0:
                print("Saindo...")
                return
            elif escolha == 99:
                caminho = input("Informe o caminho do arquivo .csv ou .xlsx: ").strip()
                padronizar_excel_agents(caminho)
                return
            cliente = list(CLIENTES.keys())[escolha - 1]
        except:
            print("[\033[1;31m!\033[m] Entrada inválida.")
            return
    else:
        cliente = args.cliente

    if cliente not in CLIENTES:
        print(f"[\033[1;31m!\033[m] Cliente '{cliente}' não configurado.")
        return

    access_key = CLIENTES[cliente]['access_key']
    secret_key = CLIENTES[cliente]['secret_key']

    if not access_key or not secret_key:
        print(f"[\033[1;31m!\033[m] Chaves de API não configuradas para o cliente {cliente}. Verifique o .env.")
        return

    output_folder = os.path.join(os.getcwd(), cliente)
    os.makedirs(output_folder, exist_ok=True)

    tio = TenableIO(access_key, secret_key)
    if not validar_credenciais(tio):
        print("[\033[1;31m!\033[m] Falha na autenticação com a API.")
        return

    try:
        agentes = list(tio.agents.list())
        print(f"\n[\033[1;32m✓\033[m] Total de agentes disponíveis na console: \033[1m{len(agentes)}\033[0m")
    except Exception as e:
        print(f"[\033[1;31m!\033[m] Erro ao recuperar agentes: {e}")
        agentes = []

    try:
        assets = list(tio.exports.assets())
        print(f"[\033[1;32m✓\033[m] Total de assets disponíveis na console: \033[1m{len(assets)}\033[0m")
    except Exception as e:
        print(f"[\033[1;31m!\033[m] Erro ao recuperar assets: {e}")
        assets = []

    tipo = args.tipo or input("\n[1] Exportar Assets\n[2] Exportar Agents\nEscolha: ").strip()

    if tipo == '1' or tipo == 'assets':
        exportar_assets(tio, output_folder, cliente)
    elif tipo == '2' or tipo == 'agents':
        filtro = args.filtro
        if not filtro:
            print("\nEscolha o filtro para agentes:")
            print("[1] Offline")
            print("[2] Sem grupo")
            print("[3] Todos")
            print("[4] Offline + Sem grupo (comparar)")
            escolha_filtro = input("Escolha (1-4): ").strip()
            filtros_map = {
                '1': 'offline',
                '2': 'nogroup',
                '3': 'all',
                '4': 'compare'
            }
            filtro = filtros_map.get(escolha_filtro, 'offline')
        exportar_agents(tio, output_folder, cliente, filtro=filtro)
    else:
        print("[\033[1;31m!\033[m] Tipo de exportação inválido.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[\033[1;31m!\033[m] Interrompido pelo usuário.")
    except Exception as e:
        print(f"[\033[1;31m!\033[m] Erro inesperado: {e}")
