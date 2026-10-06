# Currículo em LaTeX e Word

Versão editável de `Kleber_Daniel_Mattos_Viana_.pdf`, preservando o texto, a foto, os links e a organização em duas páginas.

## Arquivos

- [Fonte LaTeX](Kleber_Daniel_Mattos_Viana.tex)
- [PDF compilado](output/pdf/Kleber_Daniel_Mattos_Viana.pdf)
- [Currículo Word editável](output/word/Kleber_Daniel_Mattos_Viana.docx)
- [Pacote LaTeX para Overleaf](Kleber_Daniel_Mattos_Viana_LaTeX.zip)

## Word

A versão `.docx` pode ser aberta e editada no Microsoft Word ou LibreOffice. Contém texto editável, foto original, links clicáveis e duas páginas. O layout foi reconstruído para Word e pode apresentar pequenas diferenças em relação ao PDF.

Para manter a tipografia, instale as fontes Source Sans 3 incluídas em `assets/fonts/`. Sem essas fontes, o editor pode substituí-las e alterar as quebras de linha ou de página.

Para gerar novamente o Word após editar o LaTeX, execute na pasta do projeto:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install python-docx Pillow
.venv/bin/python scripts/exportar_word.py
```

O script lê `Kleber_Daniel_Mattos_Viana.tex`, usa `assets/foto.jpg` e grava `output/word/Kleber_Daniel_Mattos_Viana.docx`. Ele reconhece os comandos usados neste currículo; mudanças nos dados de contato ou na estrutura do cabeçalho devem ser aplicadas também ao script. A exportação substitui o Word existente; alterações feitas diretamente no `.docx` não retornam ao LaTeX.

## Compilar LaTeX

Na pasta deste projeto, com LuaLaTeX instalado:

No Ubuntu/Debian, o suporte completo de fontes exige o pacote `texlive-luatex`, além dos pacotes LaTeX usados no documento (`geometry`, `fontspec`, `babel`, `graphicx`, `tikz`, `enumitem` e `hyperref`).

```bash
mkdir -p output/pdf
lualatex -interaction=nonstopmode -halt-on-error -output-directory=output/pdf Kleber_Daniel_Mattos_Viana.tex
lualatex -interaction=nonstopmode -halt-on-error -output-directory=output/pdf Kleber_Daniel_Mattos_Viana.tex
```

O PDF resultante fica em `output/pdf/Kleber_Daniel_Mattos_Viana.pdf`.

Para usar no Overleaf, envie o `.tex` e a pasta `assets/`, selecione o arquivo principal e escolha **LuaLaTeX** como compilador. A pasta `assets/` contém a foto extraída do PDF e as fontes locais; mantenha os caminhos relativos.

Fonte Source Sans 3, sucessora de Source Sans Pro usada no original. Ícones Font Awesome. Licenças incluídas em `assets/fonts/`. A composição foi reconstruída em LaTeX, portanto pequenas diferenças tipográficas em relação ao FlowCV são esperadas.
