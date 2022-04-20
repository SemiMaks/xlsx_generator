# Генератор таблиц

from flask import Flask, render_template
from flask_sslify import SSLify
from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField

app = Flask(__name__, template_folder='.')
app.config['SECRET_KEY']='LongAndRandomSecretKey'
sslify = SSLify(app)

class TableForm(FlaskForm):
    username = StringField(label=('Enter Your Name: '))
    year = StringField(label=('Введите год: '))
    month = StringField(label=('Введите месяц: '))
    score_day = StringField(label=('Количество дней в месяце:  '))
    start_day = StringField(label=('Число первого рабочего дня: '))
    submit = SubmitField(label=('Сгенерировать'))

@app.route('/', methods=('GET', 'POST'))
def index():
    form = TableForm()
    if form.validate_on_submit():
        return f'''<h1> Добро пожаловать {form.username.data} </h1>'''
    return render_template('index.html', form=form)
    # return '<h1> Генератор таблиц </h1>'


if __name__ == '__main__':
    app.run()

# Устанавливаем библиотеку pip install flask-sslify
# для создания защищенного соединения
