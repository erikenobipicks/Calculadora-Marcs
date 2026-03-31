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

## Imatges de motllures
- Pots indicar una URL d'imatge des de l'admin del cataleg.
- També pots pujar un fitxer directament des de la fitxa de la motllura.
- Si copies una imatge dins `static/fotos/` amb el nom de la referencia, la calculadora la detecta automaticament.
  Exemple: `static/fotos/m0970.jpg`

L'importacio des de l'Excel conserva la foto ja associada a cada motllura.

## Variables d'entorn
- `SECRET_KEY` - clau secreta Flask obligatoria en produccio
- `PORT` - port de desplegament
- `ADMIN_USERNAME` - usuari del primer administrador
- `ADMIN_PASSWORD` - contrasenya del primer administrador
- `ADMIN_NAME` - nom visible del primer administrador (opcional)

L'aplicacio ja no crea un usuari `admin/admin123` per defecte. En un desplegament nou, defineix `ADMIN_USERNAME` i `ADMIN_PASSWORD` abans del primer arrenc.
