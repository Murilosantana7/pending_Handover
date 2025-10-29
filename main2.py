## quando for voltar a usar, apagar aqui import asyncio
from playwright.async_api import async_playwright
from datetime import datetime
import os
import shutil
import gspread
import gspread.utils  # Importação necessária
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

DOWNLOAD_DIR = "/tmp"

# ==============================
# Funções de renomear arquivos
# ==============================
def rename_downloaded_file(download_dir, download_path):
    try:
        current_hour = datetime.now().strftime("%H")
        new_file_name = f"PEND-{current_hour}.csv"
        new_file_path = os.path.join(download_dir, new_file_name)
        if os.path.exists(new_file_path):
            os.remove(new_file_path)
        shutil.move(download_path, new_file_path)
        print(f"Arquivo salvo como: {new_file_path}")
        return new_file_path
    except Exception as e:
        print(f"Erro ao renomear o arquivo: {e}")
        return None

# ==============================
# Funções de renomear arquivos
# ==============================
def rename_downloaded_file2(download_dir, download_path2):
    try:
        current_hour = datetime.now().strftime("%H")
        new_file_name2 = f"PROD-{current_hour}.csv"
        new_file_path2 = os.path.join(download_dir, new_file_name2)
        if os.path.exists(new_file_path2):
            os.remove(new_file_path2)
        shutil.move(download_path2, new_file_path2)
        print(f"Arquivo salvo como: {new_file_path2}")
        return new_file_path2
    except Exception as e:
        print(f"Erro ao renomear o arquivo: {e}")
        return None


# ==============================
# Funções de atualização Google Sheets (SEM PISCAR)
# ==============================
def update_packing_google_sheets(csv_file_path):
    try:
        if not os.path.exists(csv_file_path):
            print(f"Arquivo {csv_file_path} não encontrado.")
            return
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("hxh.json", scope)
        client = gspread.authorize(creds)
        
        # --- USA O NOVO ID DA PLANILHA ---
        sheet1 = client.open_by_key(
            "1qvgVViwnLVkzLnjfWQLU3m6ce0f3lXrvg-aq2YF59v8"
        )
        worksheet1 = sheet1.worksheet("Base Pending")

        # 1. Preparar novos dados
        df = pd.read_csv(csv_file_path).fillna("")
        data_to_write = [df.columns.values.tolist()] + df.values.tolist()
        new_rows = len(data_to_write)
        new_cols = len(data_to_write[0]) if new_rows > 0 else 0

        if new_rows == 0:
            worksheet1.clear()
            print(f"Arquivo CSV {csv_file_path} está vazio. Limpando a aba 'Base Pending'.")
            return

        # 2. Obter dimensões totais da planilha
        total_rows = worksheet1.row_count
        total_cols = worksheet1.col_count

        # 3. Escrever os novos dados (sem limpar primeiro)
        # --- DESTINO ATUAL: A1 ---
        worksheet1.update(data_to_write, 'A1')

        # 4. Definir os ranges para limpar dados "fantasmas"
        ranges_to_clear = []

        if new_rows < total_rows:
            start_cell_rows = gspread.utils.rowcol_to_a1(new_rows + 1, 1)
            end_cell_rows = gspread.utils.rowcol_to_a1(total_rows, total_cols)
            ranges_to_clear.append(f"{start_cell_rows}:{end_cell_rows}")

        if new_cols < total_cols:
            start_cell_cols = gspread.utils.rowcol_to_a1(1, new_cols + 1)
            end_cell_cols = gspread.utils.rowcol_to_a1(new_rows, total_cols)
            ranges_to_clear.append(f"{start_cell_cols}:{end_cell_cols}")
        
        if ranges_to_clear:
            worksheet1.batch_clear(ranges_to_clear)

        print(f"Arquivo enviado com sucesso para a aba 'Base Pending' (sem piscar).")
    except Exception as e:
        print(f"Erro durante o processo: {e}")

# ==============================
# Funções de atualização Google Sheets (SEM PISCAR)
# ==============================
def update_packing_google_sheets2(csv_file_path2):
    try:
        if not os.path.exists(csv_file_path2):
            print(f"Arquivo {csv_file_path2} não encontrado.")
            return
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("hxh.json", scope)
        client = gspread.authorize(creds)
        
        # --- USA O NOVO ID DA PLANILHA ---
        sheet1 = client.open_by_key(
            "1qvgVViwnLVkzLnjfWQLU3m6ce0f3lXrvg-aq2YF59v8"
        )
        worksheet1 = sheet1.worksheet("Base Handedover")
        
        # 1. Preparar novos dados
        df = pd.read_csv(csv_file_path2).fillna("")
        data_to_write = [df.columns.values.tolist()] + df.values.tolist()
        new_rows = len(data_to_write)
        new_cols = len(data_to_write[0]) if new_rows > 0 else 0

        if new_rows == 0:
            worksheet1.clear()
            print(f"Arquivo CSV {csv_file_path2} está vazio. Limpando a aba 'Base Handedover'.")
            return

        # 2. Obter dimensões totais da planilha
        total_rows = worksheet1.row_count
        total_cols = worksheet1.col_count

        # 3. Escrever os novos dados (sem limpar primeiro)
        # --- DESTINO ATUAL: A1 ---
        worksheet1.update(data_to_write, 'A1')

        # 4. Definir os ranges para limpar dados "fantasmas"
        ranges_to_clear = []

        if new_rows < total_rows:
            start_cell_rows = gspread.utils.rowcol_to_a1(new_rows + 1, 1)
            end_cell_rows = gspread.utils.rowcol_to_a1(total_rows, total_cols)
            ranges_to_clear.append(f"{start_cell_rows}:{end_cell_rows}")

        if new_cols < total_cols:
            start_cell_cols = gspread.utils.rowcol_to_a1(1, new_cols + 1)
            end_cell_cols = gspread.utils.rowcol_to_a1(new_rows, total_cols)
            ranges_to_clear.append(f"{start_cell_cols}:{end_cell_cols}")
        
        if ranges_to_clear:
            worksheet1.batch_clear(ranges_to_clear)

        print(f"Arquivo enviado com sucesso para a aba 'Base Handedover' (sem piscar).")
    except Exception as e:
        print(f"Erro durante o processo: {e}")
        
# ==============================
# Fluxo principal Playwright
# ==============================
async def main():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        try:
            # LOGIN (seu código de login aqui...)
            await page.goto("https://spx.shopee.com.br/")
            await page.wait_for_selector('xpath=//*[@placeholder="Ops ID"]', timeout=15000)
            await page.locator('xpath=//*[@placeholder="Ops ID"]').fill('Ops113074')
            await page.locator('xpath=//*[@placeholder="Senha"]').fill('@Shopee123')
            await page.locator('xpath=/html/body/div[1]/div/div[2]/div/div/div[1]/div[3]/form/div/div/button').click()
            await page.wait_for_load_state("networkidle", timeout=20000) # É melhor esperar a página carregar
            
            try:
                await page.locator('.ssc-dialog-close').click(timeout=5000)
            except:
                print("Nenhum pop-up foi encontrado.")
                await page.keyboard.press("Escape")

            # ================== DOWNLOAD 1: PENDING ==================
            print("\nIniciando Download 1: Base Pending")
            await page.goto("https://spx.shopee.com.br/#/hubLinehaulTrips/trip")
            await page.wait_for_timeout(8000) 
            
            # Clicando no filtro específico para "Pending" (ajuste o seletor se necessário)
            # await page.locator('xpath=/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[3]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[2]').click()
            await page.wait_for_timeout(11000)
            await page.get_by_role("button", name="Exportar").nth(0).click()
            await page.wait_for_timeout(11000)

            await page.goto("https://spx.shopee.com.br/#/taskCenter/exportTaskCenter")
            await page.wait_for_timeout(12000)
            
            async with page.expect_download() as download_info:
                await page.get_by_role("button", name="Baixar").nth(0).click()
            
            download = await download_info.value
            download_path = os.path.join(DOWNLOAD_DIR, download.suggested_filename)
            await download.save_as(download_path)

            # Usando a função unificada
            new_file_path = rename_downloaded_file(DOWNLOAD_DIR, download_path)
            if new_file_path:
                update_packing_google_sheets(new_file_path)

            # ================== DOWNLOAD 2: HANDEDOVER ==================
            print("\nIniciando Download 2: Base Handedover")
            await page.goto("https://spx.shopee.com.br/#/hubLinehaulTrips/trip")
            await page.wait_for_timeout(8000)
            
            # Clicando no filtro específico para "Handedover" (ajuste o seletor se necessário)
            await page.locator('xpath=/html[1]/body[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/span[1]').click()
            await page.wait_for_timeout(10000)
            await page.get_by_role("button", name="Exportar").nth(0).click()
            await page.wait_for_timeout(20000)

            await page.goto("https://spx.shopee.com.br/#/taskCenter/exportTaskCenter")
            await page.wait_for_timeout(30000)

            async with page.expect_download() as download_info2: # Use uma nova variável para clareza
                # Clica no botão mais recente, que deve ser o da segunda exportação
                await page.get_by_role("button", name="Baixar").nth(0).click()

            download2 = await download_info2.value
            download_path2 = os.path.join(DOWNLOAD_DIR, download2.suggested_filename) # Correção aqui
            await download2.save_as(download_path2)
            
            # Usando a função unificada
            new_file_path2 = rename_downloaded_file2(DOWNLOAD_DIR, download_path2)
            if new_file_path2:
                update_packing_google_sheets2(new_file_path2)

            print("\n✅ Processo concluído com sucesso.")

        except Exception as e:
            print(f"❌ Erro fatal durante o processo: {e}")
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
