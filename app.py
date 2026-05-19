from flask import Flask, request, jsonify, send_file, Response
import os
import tempfile
import shutil
from functools import wraps
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# Credenciales de autenticación
AUTH_USERNAME = "admin"
AUTH_PASSWORD = "admin*666"

def check_auth(username, password):
    """Verifica las credenciales"""
    return username == AUTH_USERNAME and password == AUTH_PASSWORD

def authenticate():
    """Envina respuesta de autenticación requerida"""
    return Response(
        'Autenticación requerida. Ingresa usuario y contraseña.',
        401,
        {'WWW-Authenticate': 'Basic realm="Login Required"'})

def requires_auth(f):
    """Decorador para requerir autenticación"""
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated

# Directorio base del proyecto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INFO_DIR = os.path.join(BASE_DIR, "Info")

# Crear directorio Info si no existe
os.makedirs(INFO_DIR, exist_ok=True)

@app.route('/')
@requires_auth
def index():
    """Servir el dashboard HTML estático"""
    index_path = os.path.join(BASE_DIR, "index.html")
    if os.path.exists(index_path):
        return send_file(index_path)
    return "<h1>Dashboard no encontrado. Ejecuta generar_dashboard.py primero.</h1>", 404

@app.route('/upload', methods=['POST'])
@requires_auth
def upload_file():
    """Subir un archivo Excel al directorio Info"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not file.filename.endswith('.xlsx'):
        return jsonify({'error': 'File must be .xlsx'}), 400
    
    # Guardar el archivo en Info
    filepath = os.path.join(INFO_DIR, file.filename)
    file.save(filepath)
    
    return jsonify({'success': True, 'filename': file.filename})

@app.route('/generate', methods=['POST'])
@requires_auth
def generate_dashboard():
    """Ejecutar el script de generación del dashboard"""
    try:
        # Importar y ejecutar el generador
        import sys
        sys.path.insert(0, BASE_DIR)
        
        # Ejecutar generar_dashboard.py
        exec(open(os.path.join(BASE_DIR, 'generar_dashboard.py')).read())
        
        return jsonify({'success': True, 'message': 'Dashboard generated successfully'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files', methods=['GET'])
@requires_auth
def list_files():
    """Listar archivos Excel en el directorio Info"""
    try:
        archivos = [f for f in os.listdir(INFO_DIR) if f.endswith('.xlsx') and not f.startswith('~$')]
        return jsonify({'files': archivos})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/files/<filename>', methods=['DELETE'])
@requires_auth
def delete_file(filename):
    """Eliminar un archivo Excel del directorio Info"""
    try:
        filepath = os.path.join(INFO_DIR, filename)
        if os.path.exists(filepath):
            os.remove(filepath)
            return jsonify({'success': True, 'filename': filename})
        return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# Vercel entrypoint
app = app

if __name__ == '__main__':
    app.run(debug=True, port=5000)
