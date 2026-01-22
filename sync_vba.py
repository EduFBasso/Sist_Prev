"""
Script de Sincronização VBA - Sistema Previdenciário
Sincroniza códigos VBA entre arquivos .bas e erp_prev.xlsm

Uso:
    python sync_vba.py export                    # Extrai VBA do .xlsm para arquivos .bas
    python sync_vba.py import                    # Importa TODOS os .bas para o .xlsm
    python sync_vba.py import frmSimulacoes      # Importa apenas frmSimulacoes.bas
    python sync_vba.py import modSimulacoes modUtil  # Importa múltiplos arquivos
    python sync_vba.py check                     # Verifica diferenças sem modificar

Exemplos:
    # Atualizar apenas o formulário de simulações
    python sync_vba.py import frmSimulacoes
    
    # Atualizar apenas módulos de cálculo
    python sync_vba.py import modSimulacoes modVinculos
    
    # Atualizar tudo
    python sync_vba.py import
"""

import os
import sys
import zipfile
import shutil
from pathlib import Path
import tempfile
import re

class VBASync:
    def __init__(self, xlsm_path="erp_prev.xlsm"):
        self.xlsm_path = Path(xlsm_path)
        self.base_dir = Path(__file__).parent
        self.formularios_dir = self.base_dir / "Formularios"
        self.modulos_dir = self.base_dir / "Modulos"
        self.planilhas_dir = self.base_dir / "Planilhas"
        
    def export_from_xlsm(self):
        """Extrai códigos VBA do .xlsm para arquivos .bas"""
        print(f"📤 Exportando VBA de {self.xlsm_path.name}...")
        
        if not self.xlsm_path.exists():
            print(f"❌ Erro: Arquivo {self.xlsm_path} não encontrado!")
            return False
        
        # XLSM é um arquivo ZIP
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            # Extrair o XLSM
            print("   Descompactando arquivo...")
            with zipfile.ZipFile(self.xlsm_path, 'r') as zip_ref:
                zip_ref.extractall(temp_path)
            
            # Procurar arquivos VBA
            vba_dir = temp_path / "xl" / "vbaProject.bin"
            
            if not vba_dir.parent.exists():
                print("❌ Não foi possível encontrar o projeto VBA no arquivo")
                return False
            
            # Alternativa: usar win32com para exportar (melhor opção)
            return self._export_with_win32com()
    
    def _export_with_win32com(self):
        """Exporta VBA usando win32com (requer pywin32)"""
        try:
            import win32com.client
        except ImportError:
            print("❌ Pacote 'pywin32' não instalado.")
            print("   Instale com: pip install pywin32")
            return False
        
        try:
            # Abrir Excel
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Abrir workbook
            wb = excel.Workbooks.Open(str(self.xlsm_path.absolute()))
            vb_project = wb.VBProject
            
            exported = {"forms": 0, "modules": 0, "sheets": 0}
            
            # Exportar cada componente
            for component in vb_project.VBComponents:
                comp_type = component.Type
                comp_name = component.Name
                
                # 1 = Module, 2 = Class, 3 = UserForm, 100 = Document
                if comp_type == 1:  # Módulo padrão
                    target_dir = self.modulos_dir
                    exported["modules"] += 1
                elif comp_type == 3:  # UserForm
                    target_dir = self.formularios_dir
                    exported["forms"] += 1
                elif comp_type == 100:  # Sheet/Worksheet
                    target_dir = self.planilhas_dir
                    exported["sheets"] += 1
                else:
                    continue
                
                # Criar diretório se não existir
                target_dir.mkdir(exist_ok=True)
                
                # Exportar
                file_path = target_dir / f"{comp_name}.bas"
                component.Export(str(file_path.absolute()))
                print(f"   ✓ {comp_name}.bas")
            
            # Fechar
            wb.Close(SaveChanges=False)
            excel.Quit()
            
            print(f"\n✅ Exportação concluída:")
            print(f"   • {exported['forms']} formulários")
            print(f"   • {exported['modules']} módulos")
            print(f"   • {exported['sheets']} planilhas")
            
            return True
            
        except Exception as e:
            print(f"❌ Erro ao exportar: {e}")
            return False
    
    def import_to_xlsm(self, specific_files=None):
        """
        Importa arquivos .bas para o .xlsm
        
        Args:
            specific_files: Lista de nomes de arquivos (sem extensão) para importar.
                          Se None, importa todos os arquivos.
                          Ex: ['frmSimulacoes', 'modSimulacoes']
        """
        if specific_files:
            print(f"📥 Importando módulos específicos para {self.xlsm_path.name}...")
            print(f"   Alvos: {', '.join(specific_files)}")
        else:
            print(f"📥 Importando TODOS os VBA para {self.xlsm_path.name}...")
        
        try:
            import win32com.client
        except ImportError:
            print("❌ Pacote 'pywin32' não instalado.")
            print("   Instale com: pip install pywin32")
            return False
        
        if not self.xlsm_path.exists():
            print(f"❌ Erro: Arquivo {self.xlsm_path} não encontrado!")
            return False
        
        try:
            # Abrir Excel
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            # Abrir workbook
            wb = excel.Workbooks.Open(str(self.xlsm_path.absolute()))
            vb_project = wb.VBProject
            
            imported = {"forms": 0, "modules": 0, "sheets": 0}
            
            # Coletar arquivos .bas (todos ou específicos)
            bas_files = []
            
            for directory, label in [
                (self.formularios_dir, "forms"),
                (self.modulos_dir, "modules"),
                (self.planilhas_dir, "sheets")
            ]:
                if directory.exists():
                    for bas_file in directory.glob("*.bas"):
                        # Se specific_files foi fornecido, filtrar
                        if specific_files:
                            file_name = bas_file.stem  # Nome sem extensão
                            if file_name not in specific_files:
                                continue
                        bas_files.append((bas_file, label))
            
            # Verificar se encontrou os arquivos solicitados
            if specific_files:
                found_names = [f.stem for f, _ in bas_files]
                not_found = [name for name in specific_files if name not in found_names]
                if not_found:
                    print(f"\n⚠️  Arquivos não encontrados: {', '.join(not_found)}")
                    print("   Verifique se os nomes estão corretos e se os arquivos existem.")
                    if not bas_files:
                        print("\n❌ Nenhum arquivo válido para importar.")
                        wb.Close(SaveChanges=False)
                        excel.Quit()
                        return False
            
            if not bas_files:
                print("❌ Nenhum arquivo .bas encontrado para importar.")
                wb.Close(SaveChanges=False)
                excel.Quit()
                return False
            
            # Remover apenas os componentes que serão substituídos
            print("   Removendo componentes a serem substituídos...")
            components_to_remove = []
            files_to_import = [f.stem for f, _ in bas_files]
            
            for component in vb_project.VBComponents:
                # Não remover sheets do documento (Type 100)
                if component.Type != 100:
                    # Se specific_files foi fornecido, remover apenas os especificados
                    if specific_files:
                        if component.Name in files_to_import:
                            components_to_remove.append(component.Name)
                    else:
                        # Se importando tudo, remover todos exceto sheets
                        components_to_remove.append(component.Name)
            
            for comp_name in components_to_remove:
                try:
                    vb_project.VBComponents.Remove(vb_project.VBComponents(comp_name))
                    print(f"   🗑️  Removido: {comp_name}")
                except Exception as e:
                    print(f"   ⚠️  Não foi possível remover {comp_name}: {e}")
            
            # Importar novos componentes
            print("   Importando componentes...")
            for bas_file, label in bas_files:
                vb_project.VBComponents.Import(str(bas_file.absolute()))
                imported[label] += 1
                print(f"   ✓ {bas_file.name}")
            
            # Salvar e fechar
            wb.Save()
            wb.Close(SaveChanges=True)
            excel.Quit()
            
            print(f"\n✅ Importação concluída:")
            print(f"   • {imported['forms']} formulários")
            print(f"   • {imported['modules']} módulos")
            print(f"   • {imported['sheets']} planilhas")
            
            return True
            
        except Exception as e:
            print(f"❌ Erro ao importar: {e}")
            try:
                wb.Close(SaveChanges=False)
                excel.Quit()
            except:
                pass
            return False
    
    def check_differences(self):
        """Verifica diferenças entre .bas e .xlsm"""
        print(f"🔍 Verificando diferenças...\n")
        
        # Primeiro exporta para um diretório temporário
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_xlsm = Path(temp_dir) / self.xlsm_path.name
            shutil.copy2(self.xlsm_path, temp_xlsm)
            
            # Exportar do XLSM
            print("Essa funcionalidade está em desenvolvimento.")
            print("Use 'export' para extrair e comparar manualmente.")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    
    command = sys.argv[1].lower()
    sync = VBASync()
    
    if command == "export":
        success = sync.export_from_xlsm()
    elif command == "import":
        # Verificar se há arquivos específicos nos argumentos
        specific_files = sys.argv[2:] if len(sys.argv) > 2 else None
        success = sync.import_to_xlsm(specific_files)
    elif command == "check":
        success = sync.check_differences()
    else:
        print(f"❌ Comando desconhecido: {command}")
        print(__doc__)
        sys.exit(1)
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
