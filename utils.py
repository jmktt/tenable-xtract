import pandas as pd
from datetime import datetime

def formatar_excel(df, output_file, aba):
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=aba, index=False)
            workbook = writer.book
            worksheet = writer.sheets[aba]

            (max_row, max_col) = df.shape
            column_settings = [{'header': col} for col in df.columns]

            if max_row > 0:
                worksheet.add_table(0, 0, max_row, max_col - 1, {
                    'columns': column_settings,
                    'style': 'Table Style Medium 2'
                })

            for i, col in enumerate(df.columns):
                column_len = max(df[col].astype(str).map(len).max(), len(col))
                worksheet.set_column(i, i, column_len * 1.2)
    except Exception as e:
        fallback = output_file.replace('.xlsx', '.csv')
        df.to_csv(fallback, index=False)
        print(f"[!] Erro ao salvar Excel. CSV criado: {fallback}")

def converter_timestamp(ts):
    try:
        if pd.isna(ts) or ts == 0:
            return ''
        return datetime.fromtimestamp(int(ts)).strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return ''
