================================================================================
UNIFICADOR DE PLANILHAS - IMPÉRIO DO EXTINTOR
================================================================================

COMO USAR:

1. INSTALAÇÃO (feita apenas uma vez):
   - Mantenha esta pasta no local desejado (ex: C:\Imperio do Extintor\).
   - Ela contém:
       • unificar.exe (aplicativo principal)
       • BASE.xlsm (estrutura base)
       • MODELO.xlsm (modelo com macros e botões)
       • Instalar_Menu_Contexto.bat
       • Desinstalar_Menu_Contexto.bat
   - Dê um duplo clique em "Instalar_Menu_Contexto.bat".
   - Isso adicionará a opção "Unificar" no botão direito do Windows
     e também no menu "Enviar para -> Unificar Planilhas".

2. UNIFICANDO DUAS PLANILHAS:
   - No Windows Explorer, selecione os 2 arquivos que deseja unir
     (exemplo: segure CTRL e clique em "Glaicy.xlsm" e "Taissa.xlsm").
   - Clique com o BOTÃO DIREITO em qualquer um dos dois arquivos selecionados.
   - Clique em: "Unificar" (ou "Enviar para" -> "Unificar Planilhas").
   - O aplicativo irá:
       a) Fazer o merge inteligente dos dados baseado na BASE.xlsm.
       b) Criar automaticamente cópias de segurança (.bak) de ambos os arquivos.
       c) Substituir os dois arquivos pela versão unificada e atualizada.
       d) Se não houver conflitos: exibe uma mensagem de sucesso confirmando a unificação.
       e) Se houver conflitos: exibe uma janela de alerta com a lista de conflitos e
          salva o relatório detalhado "conflitos.txt" na pasta dos arquivos.

3. USO VIA INTERFACE GRÁFICA:
   - Se preferir, dê um duplo clique diretamente em "unificar.exe".
   - Uma tela se abrirá permitindo que você escolha os 2 arquivos manualmente.

4. SE PRECISAR REMOVER DO BOTÃO DIREITO:
   - Dê um duplo clique em "Desinstalar_Menu_Contexto.bat".

================================================================================
IMPORTANTE:
- Certifique-se de que os arquivos que você vai unificar não estejam abertos
  no Excel no momento da unificação.
- Mantenha os arquivos BASE.xlsm e MODELO.xlsm na mesma pasta do unificar.exe.
================================================================================
