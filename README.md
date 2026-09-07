# Inventário do Grupo

App de contagem de inventário dos restaurantes (Adega do Leo, Cais 56, Moscatel), usada pelos gerentes no telemóvel.

- `index.html` — a app. Publicada em https://groupmultifund.github.io/inventario/
- `catalogo.json` — lista de artigos por armazém e secção (a mesma que está embutida na app).
- `Codigo.gs` — o Apps Script que recebe as contagens e as escreve na folha "INVENTÁRIO — Mestre", na pasta `02 Contagens` do Drive.

O gerente escolhe armazém e secção, conta (aceita parcelas como `10+12`) e, ao fechar, a contagem entra na folha do Drive.
Sem login: a app usa um código de acesso guardado no telemóvel.
