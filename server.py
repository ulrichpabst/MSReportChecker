from flask import Flask, send_from_directory
import ssl
import os

app = Flask(__name__, static_folder=".")

@app.route('/src/<path:filename>')
def serve_src_files(filename):
    return send_from_directory("src", filename)

@app.route('/assets/<path:filename>')
def serve_assets(filename):
    return send_from_directory("assets", filename)

@app.route('/<path:filename>')
def serve_root_files(filename):
    if os.path.exists(filename):
        return send_from_directory(".", filename)
    else:
        return "File Not Found", 404

if __name__ == '__main__':
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain(certfile="cert.pem", keyfile="key.pem")
    app.run(host="localhost", port=3000, ssl_context=context)
