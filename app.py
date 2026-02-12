from flask import Flask
from routes.ingreso import ingreso_bp
from routes.admin import admin_bp
from routes.visitantes import visitantes_bp

app = Flask(__name__)

# Registrar los Blueprints (Módulos)
app.register_blueprint(ingreso_bp)
app.register_blueprint(admin_bp)
app.register_blueprint(visitantes_bp)

if __name__ == '__main__':
    # Configura tu IP y Puerto aquí
    app.run(debug=True, host='172.16.2.169', port=5000)