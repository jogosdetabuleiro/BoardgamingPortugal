from flask import Flask, render_template_string, request
import pandas as pd

app = Flask(__name__)

# Carregar o Excel com múltiplas sheets e armazená-lo
excel_file = "Boardgamers em Portugal_BD.xlsx"
sheets_dict = pd.read_excel(excel_file, sheet_name=None)

# HTML Template com filtros e links clicáveis
html_template = """
<!doctype html>
<html lang="pt">
<head>
  <meta charset="utf-8">
  <title>Board Gamers Portugal</title>
  <style>
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 8px; text-align: left; border: 1px solid #ddd; }
    th { background-color: #f2f2f2; }
    a { color: blue; text-decoration: underline; }
  </style>
</head>
<body>
  <h1>Board Gamers Portugal</h1>

  <!-- Selecionar Sheet -->
  <form method="get">
    <label for="sheet">Escolha o tópico:</label>
    <select name="sheet" id="sheet" onchange="this.form.submit()">
      {% for sheet_name in sheet_names %}
        <option value="{{ sheet_name }}" {% if sheet_name == current_sheet %}selected{% endif %}>{{ sheet_name }}</option>
      {% endfor %}
    </select>
  </form>

  <!-- Formulário de Filtro -->
  <form method="get">
    <input type="hidden" name="sheet" value="{{ current_sheet }}">
    <label for="column">Filtrar por:</label>
    <select name="column" id="column">
      <option value="">--Selecione uma Coluna--</option>
      {% for col in columns %}
        <option value="{{ col }}" {% if col == filter_column %}selected{% endif %}>{{ col }}</option>
      {% endfor %}
    </select>
    <label for="filter_value">Valor do Filtro:</label>
    <input type="text" name="filter_value" value="{{ filter_value }}">
    <button type="submit">Aplicar Filtro</button>
  </form>

  <!-- Exibir Tabela -->
  <table>
    <thead>
      <tr>
        {% for col in columns %}
          <th>{{ col }}</th>
        {% endfor %}
      </tr>
    </thead>
    <tbody>
      {% for row in data %}
        <tr>
          {% for cell in row %}
            <td>
              {% if cell is string and (cell.startswith("http") or "maps.google.com" in cell) %}
                <a href="{{ cell }}" target="_blank">{{ cell }}</a>
              {% else %}
                {{ cell }}
              {% endif %}
            </td>
          {% endfor %}
        </tr>
      {% endfor %}
    </tbody>
  </table>
</body>
</html>
"""

@app.route("/")
def index():
    # Pega o nome da sheet selecionada (ou a primeira por padrão)
    sheet_name = request.args.get("sheet", list(sheets_dict.keys())[0])
    df = sheets_dict[sheet_name].fillna("")  # Substitui NaN por string vazia

    # Pega os valores de filtro selecionados
    filter_column = request.args.get("column", "")
    filter_value = request.args.get("filter_value", "")

    # Aplica o filtro, se um valor de filtro foi fornecido
    if filter_column and filter_value:
        df = df[df[filter_column].astype(str).str.contains(filter_value, case=False, na=False)]

    # Prepara os dados para exibição no HTML
    columns = df.columns.tolist()
    data = df.values.tolist()
    sheet_names = sheets_dict.keys()
    
    return render_template_string(
        html_template,
        columns=columns,
        data=data,
        sheet_names=sheet_names,
        current_sheet=sheet_name,
        filter_column=filter_column,
        filter_value=filter_value
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
