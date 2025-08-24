import os
import re
from datetime import datetime

import arrow
import pandas as pd

from utils import formatar_aba, TenableIO


def exportar_patch_tuesday(tio: TenableIO, output_folder: str, cliente: str):
    """Exporta vulnerabilidades do Patch Tuesday para um arquivo Excel.

    A função consulta a API de exportação de vulnerabilidades da Tenable filtrando
    pelo plugin 93962 e registros vistos nos últimos 30 dias. O campo ``output``
    é analisado para extrair a informação "Latest effective update level" e o
    campo ``last_seen`` é formatado para o padrão ``dd/MM/YYYY``. São geradas duas
    abas no Excel: ``Patch_Tuesday`` com todos os registros e ``Linux_Vulns`` com
    registros cujo sistema operacional do asset corresponde a Linux.
    """
    try:
        filtros = {
            'plugin_id': [93962],
            'last_seen': int(arrow.now().shift(days=-30).timestamp())
        }
        vulns = list(tio.exports.vulns(**filtros))

        if not vulns:
            print("[\033[1;33m!\033[m] Nenhuma vulnerabilidade encontrada para Patch Tuesday.")
            return

        df = pd.json_normalize(vulns, sep='.')

        if 'output' in df.columns:
            df['Latest effective update level'] = df['output'].str.extract(
                r'Latest\s+effective\s+update\s+level\s*:?\s*(.*)',
                expand=False
            )
        else:
            df['Latest effective update level'] = ''

        if 'last_seen' in df.columns:
            df['last_seen'] = df['last_seen'].apply(
                lambda x: datetime.fromtimestamp(int(x)).strftime('%d/%m/%Y')
                if pd.notna(x) and str(x).isdigit() else x
            )

        linux_df = df[df['asset.operating_system'].str.contains(
            r'linux', flags=re.IGNORECASE, na=False
        )].copy()

        output_file = os.path.join(output_folder, f'Tenable_patch_tuesday_{cliente}.xlsx')
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Patch_Tuesday', index=False)
            formatar_aba(writer, df, 'Patch_Tuesday')

            linux_df.to_excel(writer, sheet_name='Linux_Vulns', index=False)
            formatar_aba(writer, linux_df, 'Linux_Vulns')

        print(f"[\033[1;32m✓\033[m] Exportação Patch Tuesday concluída: {output_file}")
    except Exception as e:
        import traceback
        print(f"[\033[1;31m!\033[m] Erro ao exportar Patch Tuesday: {e}")
        traceback.print_exc()
