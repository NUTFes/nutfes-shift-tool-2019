# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, Markup
from main import register_absences
from config import PASSWORD

app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        value = request.form.get('password')
        error = ''
        messages = ''
        if value and value == PASSWORD:
            try:
                messages = register_absences()
                message = "実行しました！"
            except ValueError as error:
                message = '失敗しました.'
        elif not value:
            message = 'パスワードを入力してください.'
        elif value != PASSWORD:
            message = 'パスワードが違います.'
        else:
            message = 'エラー.'
        return render_template('form.html', messages=messages, message=message, error=error)
    else:
        return render_template('form.html')


@app.template_filter('textbr')
def textbr(arg):
    return Markup(arg.replace('\n', '<br>'))

if __name__ == '__main__':
    app.debug = True
    app.run(host='0.0.0.0', port=3800)
