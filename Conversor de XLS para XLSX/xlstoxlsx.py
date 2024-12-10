import os
import win32com.client as win32

# Função para converter arquivos .xls para .xlsx
def convert_xls_to_xlsx(input_folder, output_folder):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False  # Garante que o Excel não seja aberto visivelmente

    for root, _, files in os.walk(input_folder):
        for filename in files:
            if filename.endswith('.xls'):
                input_path = os.path.join(root, filename)
                output_filename = os.path.splitext(filename)[0] + '.xlsx'
                output_path = os.path.join(output_folder, output_filename)
                
                try:
                    wb = excel.Workbooks.Open(input_path)
                    wb.SaveAs(output_path, FileFormat=51)  # 51 é o número para o formato .xlsx
                    wb.Close(SaveChanges=False)
                    print(f"Convertido: {input_path} -> {output_path}")
                except Exception as e:
                    print(f"Erro ao converter {input_path}: {str(e)}")

    excel.Quit()

if __name__ == "__main__":
    input_folder ="Entrada.xls"  # Caminho da pasta de entrada
    output_folder = "Saída.xlsx"  # Caminho da pasta de saída

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    convert_xls_to_xlsx(input_folder, output_folder)
