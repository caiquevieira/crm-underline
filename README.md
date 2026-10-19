# CRM

CRM de leads com funil, aba Hoje, gráficos e serviços vendidos (Landing Page, Tráfego Pago, Social Media…).
Funciona de dois jeitos, com o mesmo `index.html`:

| Onde abre | Onde ficam os dados |
|---|---|
| GitHub Pages (celular ou PC, de qualquer lugar) | Planilha Google, via Apps Script (`apps-script/Codigo.gs`) |
| `iniciar.bat` no PC (http://127.0.0.1:8765) | `crm.xlsx` na pasta, via `server.py` (veja `LEIAME.txt`) |

Os dados **não** ficam neste repositório: ele só tem o programa.

## Instalar a Planilha Google (uma vez)

1. No Google Drive, abra a planilha do CRM (convertida do `crm.xlsx`, em formato Google Planilhas — não .xlsx).
2. **Extensões > Apps Script**. Apague o que houver em `Código.gs`, cole o conteúdo de `apps-script/Codigo.gs` e salve.
3. No topo do editor, escolha a função **`configurarChave`** e clique em **Executar**. Autorize o acesso à planilha.
   Em **Registro de execução** aparece a **chave de acesso**: guarde-a (gerenciador de senhas).
4. **Implantar > Nova implantação** > tipo **App da Web**:
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa** (sem a chave, o app não entrega nem grava nada)

   Copie o **URL do app da Web** (termina em `/exec`).
5. Abra o CRM no GitHub Pages, cole o URL e a chave. O aparelho guarda os dois; em cada aparelho novo, digite de novo.

Para trocar a chave (ex.: perdeu o celular), rode `configurarChave` outra vez: os aparelhos antigos deixam de acessar.
Depois de mudar o `Codigo.gs`, use **Implantar > Gerenciar implantações > Editar > Nova versão** para manter o mesmo URL.

## Bom saber

- Edição feita direto na planilha é respeitada: se o CRM estiver aberto, ele avisa antes de gravar por cima.
  Depois de editar na planilha, clique em **Recarregar** no CRM.
- As regras de edição à mão são as do `LEIAME.txt` (coluna `_id`, etapas, aba `_Servicos` etc.).
- Histórico de versões da planilha (Arquivo > Histórico de versões) substitui a pasta `backups`.
- O botão **Sair** apaga a chave do aparelho.
