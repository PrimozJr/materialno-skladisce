from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
import pandas as pd
from io import BytesIO  # Za ustvarjanje datoteke v pomnilniku

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///materiali.db'
app.config['SECRET_KEY'] = 'some_secret_key'
db = SQLAlchemy(app)

# Model za material
class Material(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    location = db.Column(db.String(100), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)

# Funkcija za ustvarjanje tabele ob zagonu
def init_db():
    with app.app_context():
        db.create_all()

# Glavna stran z vsemi materiali
@app.route('/', methods=['GET', 'POST'])
def index():
    search_name = request.form.get('search_name', '')
    search_location = request.form.get('search_location', '')

    query = Material.query
    if search_name:
        query = query.filter(Material.name.like(f'%{search_name}%'))
    if search_location:
        query = query.filter(Material.location.like(f'%{search_location}%'))

    materials = query.all()
    return render_template('index.html', materials=materials)

# Funkcija za dodajanje materiala
@app.route('/add', methods=['GET', 'POST'])
def add_material():
    if request.method == 'POST':
        name = request.form['name']
        location = request.form['location']
        quantity = request.form['quantity']
        
        new_material = Material(name=name, location=location, quantity=quantity)
        
        try:
            db.session.add(new_material)
            db.session.commit()
            flash('Material uspešno dodan!', 'success')
            return redirect(url_for('index'))
        except Exception as e:
            db.session.rollback()
            flash('Napaka pri dodajanju materiala.', 'error')
            print(e)
            return redirect(url_for('add_material'))

    return render_template('add_material.html')

# Funkcija za urejanje materiala
@app.route('/edit/<int:id>', methods=['GET', 'POST'])
def edit_material(id):
    material = Material.query.get_or_404(id)

    if request.method == 'POST':
        material.name = request.form['name']
        material.location = request.form['location']
        material.quantity = request.form['quantity']
        
        try:
            db.session.commit()
            flash('Material uspešno urejen!', 'success')
            return redirect(url_for('index'))
        except Exception as e:
            db.session.rollback()
            flash('Napaka pri urejanju materiala.', 'error')
            print(e)
            return redirect(url_for('edit_material', id=id))

    return render_template('edit_material.html', material=material)

# Funkcija za brisanje materiala
@app.route('/delete/<int:id>', methods=['POST'])
def delete_material(id):
    material = Material.query.get_or_404(id)
    
    try:
        db.session.delete(material)
        db.session.commit()
        flash('Material uspešno izbrisan!', 'success')
        return redirect(url_for('index'))
    except Exception as e:
        db.session.rollback()
        flash('Napaka pri brisanju materiala.', 'error')
        print(e)
        return redirect(url_for('index'))

# Funkcija za izvoz podatkov v Excel
@app.route('/export', methods=['GET'])
def export_to_excel():
    materials = Material.query.all()
    material_list = [{
        'ID': m.id,
        'Ime': m.name,
        'Lokacija': m.location,
        'Količina': m.quantity
    } for m in materials]
    
    # Ustvarjanje DataFrame iz seznama
    df = pd.DataFrame(material_list)
    
    # Ustvarjanje Excel datoteke v pomnilniku
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Materiali')
    output.seek(0)  # Premaknemo kazalec na začetek
    
    # Pošlje datoteko uporabniku za prenos
    return send_file(output, as_attachment=True, download_name='materiali.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# Zaganjanje aplikacije in ustvarjanje baze
if __name__ == '__main__':
    init_db()  # Ustvari tabelo ob zagonu aplikacije
    app.run(debug=True)

