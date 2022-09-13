import mammoth
import os

custom_styles = """ b => b.mark
                    u => u.initialism
                    p[style-name='Heading 1'] => h1.card
                    """


bootstrap_css = '<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta2/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-BmbxuPwQa2lc/FVzBcNJ7UAyJxM6wuqIj61tLrc4wSX0szH/Ev+nYRRuWlolflfl" crossorigin="anonymous">'
bootstrap_js = '<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta2/dist/js/bootstrap.bundle.min.js" integrity="sha384-b5kHyXgcpbZJO/tY9Ul7kGkf1S0CWuKcCD38l8YkeH8z8QjE0GmW1gYU5S9FOnJ0" crossorigin="anonymous"></script>'

for _,_, arquivos in os.walk('precisaConverter'):
    for arquivo in arquivos:
        print("Arquivo: ", arquivo)
        with open(f'precisaConverter/{arquivo}', "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file, style_map = custom_styles)
            html = result.value

        edited_html = bootstrap_css + html + bootstrap_js
        novoArquivo = arquivo.split(".")[0]
        output_filename = f'{novoArquivo}.html'
        with open(output_filename, "w") as f:
            f.writelines(edited_html)

print("Fim")
