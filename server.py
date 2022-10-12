from flask import Flask, render_template
import json
import daemon
from collections import OrderedDict

app = Flask(__name__,
    static_url_path='', 
    static_folder='./public',
    template_folder='./public')

    
@app.route("/")
    
def hello():
    return render_template('index.html')

with daemon.DaemonContext():
    app.run(host='188.225.27.209', port=11111)
