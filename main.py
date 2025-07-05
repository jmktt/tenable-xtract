import os
import argparse
from tenable.io import TenableIO
from pystyle import *
import random

from config import CLIENTES, banners, banner, banner1, banner2, banner3 

from exporters import exportar_assets, exportar_agents, exportar_vulnerabilidades 
from utils import validar_credenciais, padronizar_excel_agents, inventario_software


def main():
    parser = argparse.ArgumentParser(description="Exportador de Assets, Agents, Inventário de Software e Vulnerabilidades da Tenable")
    parser.add_argument('-c', '--cliente', help='Nome do cliente', required=False)
    parser.add_argument('-t', '--tipo', choices=['assets', 'agents','inv', 'vuln'], help='Tipo de exportação', required=False)
    parser.add_argument('-f', '--filtro', choices=['offline', 'nogroup', 'all', 'compare'], help='Filtro para agentes', required=False)
    parser.add_argument('-p', '--padronizar',help='Caminho do arquivo CSV ou XLSX para padronizar a planilha Agents',required=False, metavar='ARQUIVO')

    parser.add_argument('-s', nargs='+', choices=['critical', 'high', 'medium', 'low', 'info'], help='Filtrar vulnerabilidades por severidade (ex: --severidade critical high)', required=False)
    parser.add_argument('-d', type=int, help='Filtrar vulnerabilidades encontradas nos últimos N dias', required=False)
    parser.add_argument('-e', '--exploit', action='store_true', help='Incluir dados do Exploit-DB na exportação de vulnerabilidades.')

    args = parser.parse_args()

    selected_banner = random.choice(banners)
    if selected_banner == banner:
        print(Colorate.DiagonalBackwards(Colors.blue_to_cyan, "{}".format(selected_banner), 2))
    elif selected_banner == banner1:
        print(Colorate.DiagonalBackwards(Colors.red_to_black, "{}".format(selected_banner), 2))
    elif selected_banner == banner2: 
        print(Colorate.DiagonalBackwards(Colors.red_to_black, "{}".format(selected_banner), 2)) 
    elif selected_banner == banner3:
        print(selected_banner) 
    else:
        print(selected_banner)
    print("[...]          Tenable Assets & Agents Exporter       [...]\n")

    if args.padronizar:
        padronizar_excel_agents(args.padronizar)
        return

    cliente = args.cliente
    if not cliente:
        while True: 
            box_content_lines = ["Clientes disponíveis:\n"]
            for i, c in enumerate(CLIENTES):
                box_content_lines.append(f" [{i + 1}] {c}")
            box_content_lines.append("") 
            desired_width_for_content = 41 
            padded_box_content_lines = []
            for line in box_content_lines:
                padded_line = line.ljust(desired_width_for_content)
                padded_box_content_lines.append(padded_line)
            box_content = "\n".join(padded_box_content_lines)
            print(Box.DoubleCube(box_content))
            print("  [0] Sair")
            print("\n  [99] Padronizar Excel (Agents)")

            try:
                escolha_input = input("\n\033[4mxtract\033[0m> ")
                escolha = int(escolha_input)
                
                if escolha == 0:
                    print("Saindo...")
                    return
                elif escolha == 99:
                    caminho = input("Informe o caminho do arquivo .csv ou .xlsx para padronizar: ").strip()
                    if not caminho:
                        print("[\033[1;31m!\033[m] Caminho do arquivo não pode ser vazio. Tente novamente.")
                        continue 
                    padronizar_excel_agents(caminho)
                    return 
                elif 1 <= escolha <= len(CLIENTES):
                    cliente = list(CLIENTES.keys())[escolha - 1]
                    break 
                else:
                    print("[\033[1;31m!\033[m] Escolha inválida. Por favor, selecione um número da lista.")
            except ValueError:
                print("[\033[1;31m!\033[m] Entrada inválida. Por favor, digite um número.")
            except IndexError:
                print("[\033[1;31m!\033[m] Escolha inválida. Por favor, selecione um número da lista.")
    
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

    agentes = []
    try:
        agentes = list(tio.agents.list())
        print(f"[\033[1;32m✓\033[m] Total de agentes disponíveis na console: \033[1m{len(agentes)}\033[0m")
    except Exception as e:
        print(f"[\033[1;31m!\033[m] Erro ao recuperar agentes: {e}")

    assets = []
    try:
        assets = list(tio.exports.assets())
        print(f"[\033[1;32m✓\033[m] Total de assets disponíveis na console: \033[1m{len(assets)}\033[0m")
    except Exception as e:
        print(f"[\033[1;31m!\033[m] Erro ao recuperar assets: {e}")

    tipo = args.tipo
    if not tipo:
        while True: 
            print("\n[1] Exportar Assets")
            print("[2] Exportar Agents")
            print("[3] Software Inventory")
            print("[4] Exportar Vulnerabilidades") 
            tipo_input = input("\n \033[4mxtract\033[0m> ").strip()
            
            if tipo_input == '1':
                tipo = 'assets'
                break
            elif tipo_input == '2':
                tipo = 'agents'
                break
            elif tipo_input == '3':
                tipo = 'inv'
                break
            elif tipo_input == '4':
                tipo = 'vuln'
                break
            else:
                print("[\033[1;31m!\033[m] Tipo de exportação inválido. Por favor, escolha 1, 2, 3 ou 4.")

    if tipo == 'assets':
        exportar_assets(tio, output_folder, cliente, assets_data=assets) 
    elif tipo == 'agents':
        filtro = args.filtro
        if not filtro:
            while True: 
                print("\nEscolha o filtro para agentes:")
                print("[1] Offline")
                print("[2] Sem grupo")
                print("[3] Todos")
                print("[4] Offline + Sem grupo (comparar)")
                escolha_filtro_input = input("\n \033[4mxtract\033[0m> ")
                try:
                    escolha_filtro = int(escolha_filtro_input)
                    filtros_map = {
                        1: 'offline',
                        2: 'nogroup',
                        3: 'all',
                        4: 'compare'
                    }
                    if escolha_filtro in filtros_map:
                        filtro = filtros_map[escolha_filtro]
                        break 
                    else:
                        print("[\033[1;31m!\033[m] Escolha de filtro inválida. Por favor, selecione um número de 1 a 4.")
                except ValueError:
                    print("[\033[1;31m!\033[m] Entrada inválida. Por favor, digite um número.")
        exportar_agents(tio, output_folder, cliente, filtro=filtro, agents_data=agentes) 
    elif tipo == 'inv':
        inventario_software(access_key, secret_key, output_folder, cliente)
    elif tipo == 'vuln':
        severidade = args.s
        ultimos_dias = args.d
        include_exploitdb = args.exploit

        if not severidade and not ultimos_dias and not include_exploitdb:
            print("\nPara exportar vulnerabilidades, você pode usar filtros:")
            print("  [1] Por Severidade (Critical, High, Medium, Low, Info)")
            print("  [2] Pelos Últimos N Dias")
            print("  [3] Sem filtro (últimos 90 dias padrão da API)")
            
            escolha_filtro_vuln = input("\n \033[4mxtract\033[0m> ").strip()

            if escolha_filtro_vuln == '1':
                input_severidade = input("Digite as severidades separadas por espaço (ex: critical high medium): ").lower().split()
                severidades_validas = ['critical', 'high', 'medium', 'low', 'info']
                severidade = [s for s in input_severidade if s in severidades_validas]
                if not severidade:
                    print("[\033[1;33m!\033[m] Nenhuma severidade válida informada. Exportando sem filtro de severidade.")
            elif escolha_filtro_vuln == '2':
                try:
                    ultimos_dias = int(input("Digite o número de dias: "))
                except ValueError:
                    print("[\033[1;33m!\033[m] Entrada inválida para número de dias. Exportando sem filtro de dias.")
                    ultimos_dias = None
            elif escolha_filtro_vuln == '3':
                print("Exportando todas as vulnerabilidades (até os últimos 90 dias).")
            else:
                print("[\033[1;33m!\033[m] Opção de filtro inválida. Exportando vulnerabilidades sem filtro.")

        if not args.exploit: 
            include_exploitdb_input = input("Incluir dados do Exploit-DB na exportação de vulnerabilidades? (s/N): ").strip().lower()
            if include_exploitdb_input == 's' or include_exploitdb_input == 'sim' or include_exploitdb_input == 'yes' or include_exploitdb_input == 'y':
                include_exploitdb = True
            else:
                include_exploitdb = False

        exportar_vulnerabilidades(tio, output_folder, cliente, filtro_severidade=severidade, last_found_days=ultimos_dias, include_exploitdb=include_exploitdb)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[\033[1;31m!\033[m] Interrompido pelo usuário.")
    except Exception as e:
        import traceback 
        print(f"[\033[1;31m!\033[m] Erro inesperado: {e}")
        traceback.print_exc()