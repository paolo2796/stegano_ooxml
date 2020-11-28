import cgi
import gzip
import http.server
import os
import shutil
import socketserver
import urllib
from zipfile import ZipFile

from Crypto.Random import random

from decoder import decoding
from encoder import encoding


class MyHttpRequestHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        qs = {}
        path = ""
        if self.path == '/decoder' or self.path.startswith("/decoder?testo_segreto="):
            self.path = '/app/decoder.html'
        elif self.path == '/encoder' or self.path == "/encoder?success=true":
            self.path = '/app/encoder.html'


        if '?' in self.path:
            path, tmp = self.path.split('?', 1)
            if tmp:
                qs = urllib.parse.parse_qs(tmp)
        if path == '/encoding':
            encoding(qs['input'][0],qs['message'][0],qs['pass'][0])
            self.send_response(301)
            self.send_header('Location', "/encoder?success=true")
            self.end_headers()
        return http.server.SimpleHTTPRequestHandler.do_GET(self)

    def do_POST(self):

        if self.path == '/decoding':
            # Parse the form data posted
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={'REQUEST_METHOD': 'POST',
                         'CONTENT_TYPE': self.headers['Content-Type'],
                         })
            filename = form['file_stego'].filename
            data = form['file_stego'].file.read()
            open("app/stego/stego.zip", "wb").write(data)
            # Create a ZipFile Object and load sample.zip in it
            with ZipFile('app/stego/stego.zip', 'r') as zipObj:
                # Extract all the contents of zip file in different directory
                zipObj.extractall('app/stego/file_extracted')
            os.remove("app/stego/stego.zip")
            secret_text = decoding(form.getvalue("password"))
            self.send_response(301)
            self.send_header('Location', "/decoder?testo_segreto=" + secret_text)
            self.end_headers()

        return http.server.SimpleHTTPRequestHandler.do_GET(self)

# Create an object of the above class
handler_object = MyHttpRequestHandler
rand = random.choice(range(7000, 8999))
print(rand)
PORT = rand

my_server = socketserver.TCPServer(("", PORT), handler_object)
# Star the server
my_server.serve_forever()