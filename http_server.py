# Formulário de Plano de Curso – UFAC

## Como usar

1. Execute o servidor:
   ```
   python3 http_server.py 8765
   ```
   ou simplesmente:
   ```
   bash start.sh
   ```

2. Abra no navegador: http://localhost:8765

3. Preencha o formulário e clique em "Gerar Plano de Curso (.docx)"

## Requisitos
- Python 3.8+
- python-docx (`pip install python-docx`)

## Arquivos
- `index.html` — Interface web do formulário
- `http_server.py` — Servidor HTTP local
- `generate_docx.py` — Gerador do arquivo DOCX
- `logo.png` — Logo da UFAC
- `../disciplinas.json` — Base de disciplinas exportada da planilha
