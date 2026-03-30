# Objectiu Emmarcacio - Calculadora

App de calcul de pressupostos per a emmarcacio.

## Desplegament local
```bash
pip install -r requirements.txt
python app.py
```

## Importacio de catalegs
```bash
python importar_excel.py Marcs_Objectiu_2026.xlsx
python importar_impressio.py
```

## Variables d'entorn
- `SECRET_KEY` - clau secreta Flask obligatoria en produccio
- `PORT` - port de desplegament
- `ADMIN_USERNAME` - usuari del primer administrador
- `ADMIN_PASSWORD` - contrasenya del primer administrador
- `ADMIN_NAME` - nom visible del primer administrador (opcional)

L'aplicacio ja no crea un usuari `admin/admin123` per defecte. En un desplegament nou, defineix `ADMIN_USERNAME` i `ADMIN_PASSWORD` abans del primer arrenc.
