# 🧯 Unificador de Planilhas - Império do Extintor

[![Python](https://img.shields.io/badge/Python-3.13+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-Proprietary-red.svg)](#)
[![Windows](https://img.shields.io/badge/OS-Windows%2010%20%7C%2011-0078D6.svg)](#)

Ferramenta desenvolvida para o **Império do Extintor** para unificar e sincronizar planilhas de controle (`.xlsm` / `.xlsx`) via algoritmo de **3-Way Merge**, preservando 100% de macros VBA, botões, validações de dados e desenhos/logotipos originais.

<p align="center">
  <img src="icone.png" alt="Ícone Império do Extintor" width="160" />
</p>

---

## ✨ Funcionalidades Principais

- **🔄 Algoritmo 3-Way Merge Inteligente:**
  - Identifica cada linha da tabela `BASE TROCA` pela tupla `(Nº PEDIDO, TIPO, ordem dentro do pedido)`.
  - Mescla novas linhas inseridas por qualquer um dos operadores na posição correta.
  - Detecta e preserva edições exclusivas de cada operador.
- **⚠️ Resolução Interativa de Conflitos:**
  - Quando a mesma célula ou pedido é alterado simultaneamente por ambos os operadores, abre uma **janela modal interativa** com preview detalhado em tabela (`Treeview`).
  - Permite escolher qual planilha priorizar (*Planilha 1* ou *Planilha 2*).
  - Permite **Abortar** a operação a qualquer momento, sem alterar nenhum arquivo original.
  - Botão dedicado para abrir o relatório completo (`conflitos.txt`) diretamente no Bloco de Notas.
- **⏳ Indicador Visual de Loading:**
  - Barra de progresso animada indeterminada com feedback em tempo real durante a leitura, processamento e gravação das planilhas.
- **🖱️ Integração com o Botão Direito do Windows (Menu de Contexto):**
  - Selecione 2 planilhas no Windows Explorer, clique com o botão direito e escolha **"Unificar"**.
  - Ambas as planilhas selecionadas são substituídas de forma atômica pela versão unificada e atualizada.
- **⚡ Suporte Completo ao Windows 11 (Abertura no 1º Clique):**
  - Inclui script para ativar a exibição direta no menu principal do Windows 11, sem precisar clicar em *"Mostrar mais opções"*.
- **🛡️ Segurança e Rollback Automático:**
  - Cria backups automáticos (`.bak`) antes de substituir qualquer arquivo.
  - Tratamento para arquivos abertos no Excel com opção de "Repetir".
  - Se um arquivo falhar na gravação, o outro é revertido imediatamente para o backup original.
- **🎨 Ícone Personalizado em Alta Definição:**
  - Ícone 3D estilizado incorporado ao `.exe`, à janela gráfica e ao menu do Windows Explorer.

---

## 📁 Estrutura de Arquivos

```text
Imperio do Extintor/
├── unificar.exe                  # Executável pronto para uso
├── unificar.py                   # Interface gráfica (Tkinter) e coordenação
├── merge_planilhas.py            # Motor matemático do 3-way merge e manipulação XML
├── installer.py                  # Registro e configuração no Windows Registry
├── BASE.xlsm                     # Planilha base (ancestral comum)
├── MODELO.xlsm                   # Modelo com código VBA, botões e desenhos
├── icone.ico                     # Ícone do Windows (16x16 até 256x256)
├── icone.png                     # Imagem de alta resolução do logotipo
├── Instalar_Menu_Contexto.bat    # Script para registrar a opção no botão direito
├── Desinstalar_Menu_Contexto.bat # Script para remover a opção do botão direito
├── Ativar_Menu_Direto_Windows11.bat   # Ativa o menu direto no 1º clique (Windows 11)
├── Restaurar_Menu_Padrao_Windows11.bat # Restaura o menu padrão original (Windows 11)
└── README.md
```

---

## 🚀 Como Usar

### 1. Instalação Inicial (feita uma única vez)

1. Mantenha os arquivos na pasta desejada (por exemplo, `C:\Imperio do Extintor\`).
2. Dê um duplo clique em **`Instalar_Menu_Contexto.bat`**.
3. O script perguntará se você deseja ativar o **Menu Direto no Windows 11** (para abrir no 1º clique). Digite `S` e pressione `Enter`.

### 2. Unificando pelo Botão Direito do Mouse

1. No Windows Explorer, selecione as duas planilhas que deseja combinar (segure `Ctrl` e clique nos dois arquivos, por exemplo `Glaicy.xlsm` e `Taissa.xlsm`).
2. Clique com o **botão direito** em qualquer um dos arquivos selecionados.
3. Clique em **Unificar**.
4. A tela de progresso será exibida. Se houver divergências, a tela de resolução de conflitos permitirá escolher qual planilha priorizar ou abortar.
5. Ao concluir, ambos os arquivos estarão 100% atualizados e sincronizados!

### 3. Uso via Interface Gráfica

Se preferir, dê um duplo clique diretamente em **`unificar.exe`**:
- Uma janela amigável será exibida com campos para escolher a Planilha 1 e a Planilha 2 manualmente via botão "Procurar...".

---

## 🛠️ Compilação do Executável (.exe)

Caso deseje recompilar o executável a partir do código fonte em Python:

```powershell
pip install openpyxl pillow pyinstaller
python -m PyInstaller --noconfirm --onefile --windowed --name unificar --icon=icone.ico --hidden-import openpyxl --hidden-import openpyxl.worksheet.datavalidation --hidden-import PIL --hidden-import PIL.ImageTk unificar.py
```

---

## 📜 Licença

Propriedade exclusiva do **Império do Extintor**. Todos os direitos reservados.
