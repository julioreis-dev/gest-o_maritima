from flask import Flask, render_template, jsonify, request
from flask_sqlalchemy import SQLAlchemy
from environs import Env

app = Flask(__name__)

# ___________________________ Variabel Enviroment _______________________________________ #
env = Env()
env.read_env()
username = env.str('NT_username')
password = env.str('NT_password')
host = env.str('NT_host')
server = env.int('NT_server')

# __________________________ Connect to Database Mysql ___________________________________ #
app.config['SQLALCHEMY_DATABASE_URI'] = f"mysql+mysqlconnector://{username}:{password}@{host}:{server}/bd_contratos"
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)


# __________________________ TABLE Configuration _____________________________________ #
class Contratos(db.Model):
    __tablename__ = 'Contratos'
    id = db.Column(db.Integer, primary_key=True)
    icj = db.Column(db.String(20), unique=True, nullable=False)
    equipamento = db.Column(db.String(10), unique=False, nullable=True)
    embarcacao = db.Column(db.String(50), unique=False, nullable=False)
    chave_gerente = db.Column(db.String(4), unique=False, nullable=True)
    gerente = db.Column(db.String(20), unique=False, nullable=True)
    chave_fiscal = db.Column(db.String(4), unique=False, nullable=True)
    fiscal = db.Column(db.String(20), unique=False, nullable=True)
    empresa = db.Column(db.String(100), unique=False, nullable=True)
    inicio = db.Column(db.Date())
    termino = db.Column(db.Date())
    cessao = db.Column(db.String(10), unique=False, nullable=True)

    def to_dict(self):
        return {col.name: getattr(self, col.name) for col in self.__table__.columns}


# _________________________________________ Home __________________________________________ #
@app.route('/')
def home():
    return render_template('index.html')


# _____________________________ HTTP GET - Read Record ____________________________________ #
@app.route('/vessels', methods=['GET', 'POST', 'DELETE', 'PATCH'])
def fleet():
    if request.method == 'GET':
        ships = db.session.query(Contratos).all()
        return jsonify(vessels=[ship.to_dict() for ship in ships])

    if request.method == 'POST':
        new = Contratos(
            icj=request.form.get('icj'),
            equipamento=request.form.get('equipamento'),
            embarcacao=request.form.get('embarcacao'),
            chave_gerente=request.form.get('chave_gerente'),
            gerente=request.form.get('gerente'),
            chave_fiscal=request.form.get('chave_fiscal'),
            fiscal=request.form.get('fiscal'),
            empresa=request.form.get('empresa'),
            # inicio = request.form.get('inicio'),
            # termino = request.form.get('termino'),
            cessao=request.form.get('cessao'),
        )
        db.session.add(new)
        db.session.commit()
        return jsonify(response={"success": "Successfully added the new cafe."})

    if request.method == 'DELETE':
        secret = request.args.get('validation')
        if secret == 'TopSecretAPIKey':
            _icj = request.args.get('icj')
            register = db.session.query(Contratos).filter_by(icj=_icj).first()
            if register:
                db.session.delete(register)
                db.session.commit()
                return jsonify(response={"success": "Successfully deleted the cafe from the database."}), 200
            else:
                return jsonify(error={"Not Found": "Sorry a cafe with that id was not found in the database."}), 404
        else:
            return jsonify(
                error={"Forbidden": "Sorry, that's not allowed. Make sure you have the correct api_key."}), 403

    if request.method == 'PATCH':
        vessel_icj = request.args.get('icj')
        name = request.args.get('manager')
        key_manager = request.args.get('key')
        result = db.session.query(Contratos).filter_by(icj=vessel_icj).first()
        if result:
            result.gerente = name
            result.chave_gerente = key_manager
            db.session.commit()
            return jsonify(response={"success": "Successfully updated the manager from the database."}), 200
        else:
            return jsonify(error={"Not Found": "Sorry a cafe with that id was not found in the database."}), 404


@app.route('/vessels/ship')
def get_vessel():
    num = request.args.get('num')
    if len(num) > 8:
        ship = db.session.query(Contratos).filter_by(icj=num).first_or_404()
        return jsonify(vessel=ship.to_dict())
    else:
        ship = db.session.query(Contratos).filter_by(equipamento=num).first_or_404()
        return jsonify(vessel=ship.to_dict())


@app.route('/vessels/members/key')
def get_managers_key():
    key = request.args.get('cod')
    contracts = db.session.query(Contratos).filter_by(chave_gerente=key).all()
    if len(contracts) == 0:
        new_contracts = db.session.query(Contratos).filter_by(chave_fiscal=key).all()
        return jsonify(vessel=[resp.to_dict() for resp in new_contracts])
    else:
        return jsonify(vessel=[resp.to_dict() for resp in contracts])


@app.route('/vessels/members/name')
def get_managers():
    name = request.args.get('desc')
    contracts = db.session.query(Contratos).filter_by(gerente=name).all()
    if len(contracts) == 0:
        new_contracts = db.session.query(Contratos).filter_by(fiscal=name).all()
        return jsonify(vessel=[resp.to_dict() for resp in new_contracts])
    else:
        return jsonify(vessel=[resp.to_dict() for resp in contracts])


if __name__ == "__main__":
    app.run(debug=False)
