import os
from app import app, init_db

# Inicializar la base de datos al arrancar en Render
init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
