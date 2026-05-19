from flask import Flask, request, jsonify, send_file, Response
import os
import tempfile
import shutil
from functools import wraps
from flask_cors import CORS

app = Flask(__name__)
CORS(app, supports_credentials=True)

# Credenciales de autenticación
AUTH_USERNAME = "admin"
AUTH_PASSWORD = "admin*666"

def check_auth():
    """Verifica las credenciales desde headers o query params"""
    # Verificar en headers
    username = request.headers.get('X-Auth-Username')
    password = request.headers.get('X-Auth-Password')
    if username and password:
        return username == AUTH_USERNAME and password == AUTH_PASSWORD
    
    # Verificar en query params
    username = request.args.get('username')
    password = request.args.get('password')
    if username and password:
        return username == AUTH_USERNAME and password == AUTH_PASSWORD
    
    return False

def requires_auth(f):
    """Decorador para requerir autenticación"""
    @wraps(f)
    def decorated(*args, **kwargs):
        if not check_auth():
            # Si es la ruta principal, devolver página de login HTML
            if request.path == '/':
                login_page = '''
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login - Dashboard Horas Extras</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #0f172a; color: #e2e8f0; display: flex; align-items: center; justify-content: center; min-height: 100vh; margin: 0; }
        .login-box { background: linear-gradient(135deg, #1e293b, #334155); padding: 40px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.3); text-align: center; max-width: 400px; width: 90%; }
        h1 { color: #38bdf8; margin-bottom: 20px; }
        input { background: #0f172a; color: #e2e8f0; border: 1px solid #475569; border-radius: 6px; padding: 12px; width: 100%; margin-bottom: 12px; font-size: 1rem; box-sizing: border-box; }
        button { background: #3b82f6; color: #fff; border: none; padding: 12px 24px; border-radius: 6px; cursor: pointer; font-weight: 600; font-size: 1rem; width: 100%; }
        button:hover { background: #2563eb; }
        .error { color: #ef4444; margin-top: 12px; font-size: 0.9rem; }
    </style>
</head>
<body>
    <div class="login-box">
        <h1>🔐 Login</h1>
        <input type="text" id="username" placeholder="Usuario" autofocus>
        <input type="password" id="password" placeholder="Contraseña">
        <button onclick="doLogin()">Iniciar Sesión</button>
        <div id="error" class="error"></div>
    </div>
    <script>
        function doLogin() {
            const username = document.getElementById('username').value;
            const password = document.getElementById('password').value;
            const errorDiv = document.getElementById('error');
            
            if (!username || !password) {
                errorDiv.textContent = 'Ingresa usuario y contraseña';
                return;
            }
            
            // Guardar en localStorage y recargar
            localStorage.setItem('auth', JSON.stringify({ username, password }));
            location.href = location.pathname + '?username=' + encodeURIComponent(username) + '&password=' + encodeURIComponent(password);
        }
        
        // Verificar si hay credenciales guardadas
        const saved = localStorage.getItem('auth');
        if (saved) {
            try {
                const auth = JSON.parse(saved);
                document.getElementById('username').value = auth.username;
                document.getElementById('password').value = auth.password;
            } catch (e) {}
        }
        
        // Permitir Enter
        document.getElementById('password').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') doLogin();
        });
    </script>
</body>
</html>
                '''
                return login_page, 401
            return jsonify({'error': 'Autenticación requerida', 'authenticated': False}), 401
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
