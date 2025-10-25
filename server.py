from flask import Flask, request, jsonify
import subprocess

app = Flask(__name__)

@app.route('/run-script', methods=['POST'])
def run_script():
    try:
        # 실행할 Python 스크립트 경로
        script_path = "C:/path/to/your_script.py"  
        
        # 스크립트 실행
        subprocess.run(["python", script_path], check=True)

        return jsonify({"status": "success", "message": "스크립트 실행 완료!"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
