from flask import Flask, render_template, redirect, url_for, request, send_file, jsonify
from flask_mail import Mail, Message
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_LEFT
from sqlalchemy.engine.url import URL, make_url
from reportlab.platypus import Image
import requests
import json
import random
import io
import pyexcel as p
import os

app = Flask(__name__)

uri = os.getenv("DATABASE_URL", "sqlite:///local.db")
if uri.startswith("postgres://"):
    uri = uri.replace("postgres://", "postgresql://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = uri

db = SQLAlchemy(app)

# Configurações do Flask-Mail
app.config["MAIL_SERVER"] = "smtp.hostinger.com"
app.config["MAIL_PORT"] = 465
app.config["MAIL_USE_TLS"] = False
app.config["MAIL_USE_SSL"] = True
app.config["MAIL_USERNAME"] = os.getenv("MAIL_USERNAME")
app.config["MAIL_PASSWORD"] = os.getenv("MAIL_PASSWORD")
app.config["MAIL_DEFAULT_SENDER"] = os.getenv("MAIL_DEFAULT_SENDER")
mail = Mail(app)


class Inscricao(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codigo_unico = db.Column(db.String(10), unique=True, nullable=False)
    nome_completo = db.Column(db.String(255), nullable=False)  # Nome pode ser longo
    cpf = db.Column(
        db.String(14), nullable=False
    )  # Formato com máscara: 000.000.000-00
    nome_estabelecimento = db.Column(db.String(255), nullable=False)
    volume_producao_anual = db.Column(db.Integer, nullable=False)
    cnpj = db.Column(
        db.String(18), nullable=True
    )  # Formato com máscara: 00.000.000/0000-00
    telefone = db.Column(
        db.String(20), nullable=False
    )  # Espaço para formatos internacionais +00 (00) 00000-0000
    email = db.Column(db.String(255), nullable=False)
    endereco = db.Column(db.String(255), nullable=False)
    municipio = db.Column(db.String(120), nullable=False)
    estado = db.Column(db.String(2), nullable=False)  # Sigla do estado
    cep = db.Column(db.String(10), nullable=False)  # Formato com máscara: 00000-000
    nome_produto = db.Column(db.String(255), nullable=False)
    registro_estabelecimento_mapa = db.Column(
        db.String(255), nullable=False
    )  # Pode conter códigos longos
    registro_produto_mapa = db.Column(db.String(255), nullable=False)
    categoria_inscrita = db.Column(db.String(50), nullable=False)
    pasteurizado = db.Column(db.Boolean, nullable=False)
    data_fabricacao_amostras = db.Column(db.Date, nullable=False)
    lote = db.Column(db.String(50), nullable=False)
    quantidade_unidades_amostrais = db.Column(db.Integer, nullable=False)
    embalagem_amostral = db.Column(db.String(50), nullable=False)
    quantidade_ml_amostral = db.Column(db.Integer, nullable=False)
    informacoes_adicionais = db.Column(db.Text, nullable=True)
    origem_conhecimento = db.Column(db.String(255), nullable=True)
    outro_origem_conhecimento = db.Column(db.String(255), nullable=True)
    historia_producao = db.Column(db.Text, nullable=False)
    aceitou_termos = db.Column(db.Boolean, nullable=False, default=False)
    data_hora_inscricao = db.Column(
        db.DateTime, default=datetime.utcnow, nullable=False
    )

    def to_dict(self):
        return {
            "id": self.id,
            "codigo_unico": self.codigo_unico,
            "nome_completo": self.nome_completo,
            "cpf": self.cpf,
            "nome_estabelecimento": self.nome_estabelecimento,
            "volume_producao_anual": self.volume_producao_anual,
            "cnpj": self.cnpj,
            "telefone": self.telefone,
            "email": self.email,
            "endereco": self.endereco,
            "municipio": self.municipio,
            "estado": self.estado,
            "cep": self.cep,
            "nome_produto": self.nome_produto,
            "registro_estabelecimento_mapa": self.registro_estabelecimento_mapa,
            "registro_produto_mapa": self.registro_produto_mapa,
            "categoria_inscrita": self.categoria_inscrita,
            "pasteurizado": self.pasteurizado,
            "data_fabricacao_amostras": self.data_fabricacao_amostras.strftime(
                "%Y-%m-%d"
            ),
            "lote": self.lote,
            "quantidade_unidades_amostrais": self.quantidade_unidades_amostrais,
            "embalagem_amostral": self.embalagem_amostral,
            "quantidade_ml_amostral": self.quantidade_ml_amostral,
            "informacoes_adicionais": self.informacoes_adicionais,
            "origem_conhecimento": self.origem_conhecimento,
            "outro_origem_conhecimento": self.outro_origem_conhecimento,
            "historia_producao": self.historia_producao,
            "data_hora_inscricao": self.data_hora_inscricao.strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            "aceitou_termos": self.aceitou_termos,
        }


def create_tables():
    with app.app_context():
        db.create_all()


@app.route("/", methods=["GET"])
def index():
    return render_template("formulario.html")


def send_email(to_email, subject, html_content, pdf_file=None):
    with app.app_context():
        msg = Message(subject, recipients=[to_email], html=html_content)
        if pdf_file:
            msg.attach("Inscricao.pdf", "application/pdf", pdf_file.read())
        mail.send(msg)


def gerar_codigo_unico():
    existe = True
    codigo = ""
    while existe:
        codigo = "CNA-" + "".join([str(random.randint(0, 9)) for _ in range(4)])
        existe = Inscricao.query.filter_by(codigo_unico=codigo).first() is not None
    return codigo


app.config["LOGO_PATH"] = "static/logo.png"


@app.route("/inscricao", methods=["POST"])
def add_inscricao():
    nome_completo = request.form["nome_completo"]
    cpf = request.form["cpf"]
    nome_estabelecimento = request.form["nome_estabelecimento"]
    volume_producao_anual = request.form.get("volume_producao_anual")
    cnpj = request.form.get("cnpj")
    telefone = request.form["telefone"]
    email = request.form["email"]
    endereco = request.form["endereco"]
    municipio = request.form["municipio"]
    estado = request.form["estado"]
    cep = request.form["cep"]
    nome_produto = request.form["nome_produto"]
    registro_estabelecimento_mapa = request.form["registro_estabelecimento_mapa"]
    registro_produto_mapa = request.form["registro_produto_mapa"]
    categoria_inscrita = request.form["categoria_inscrita"]
    pasteurizado = request.form["pasteurizado"] == "true"
    data_fabricacao_amostras = datetime.strptime(
        request.form["data_fabricacao_amostras"], "%Y-%m-%d"
    )
    lote = request.form["lote"]
    quantidade_unidades_amostrais = int(request.form["quantidade_unidades_amostrais"])
    embalagem_amostral = request.form["embalagem_amostral"]
    quantidade_ml_amostral = int(request.form["quantidade_ml_amostral"])
    historia_producao = request.form["historia_producao"]
    informacoes_adicionais = request.form.get("informacoes_adicionais")
    origem_conhecimento = request.form.get("origem_conhecimento")
    outro_origem_conhecimento = request.form.get("outro_origem_conhecimento")
    aceitou_termos = request.form.get("aceitou_termos") == "on"
    codigo_unico = gerar_codigo_unico()

    nova_inscricao = Inscricao(
        codigo_unico=codigo_unico,
        nome_completo=nome_completo,
        cpf=cpf,
        nome_estabelecimento=nome_estabelecimento,
        volume_producao_anual=volume_producao_anual,
        cnpj=cnpj,
        telefone=telefone,
        email=email,
        endereco=endereco,
        municipio=municipio,
        estado=estado,
        cep=cep,
        nome_produto=nome_produto,
        registro_estabelecimento_mapa=registro_estabelecimento_mapa,
        registro_produto_mapa=registro_produto_mapa,
        categoria_inscrita=categoria_inscrita,
        pasteurizado=pasteurizado,
        data_fabricacao_amostras=data_fabricacao_amostras,
        lote=lote,
        quantidade_unidades_amostrais=quantidade_unidades_amostrais,
        embalagem_amostral=embalagem_amostral,
        quantidade_ml_amostral=quantidade_ml_amostral,
        informacoes_adicionais=informacoes_adicionais,
        origem_conhecimento=origem_conhecimento,
        outro_origem_conhecimento=outro_origem_conhecimento,
        historia_producao=historia_producao,
        aceitou_termos=aceitou_termos,
        data_hora_inscricao=datetime.utcnow(),
    )

    db.session.add(nova_inscricao)
    db.session.commit()

    # Geração do PDF
    pdf_buffer = create_pdf(nova_inscricao, app.config["LOGO_PATH"])

    # Configuração e envio do email
    to_email = nova_inscricao.email  # Email do usuário
    subject = "Confirmação de Inscrição"
    html_content = render_template(
        "email_template.html", nome=nova_inscricao.nome_completo
    )

    # Enviar email com conteúdo HTML e anexo PDF (se existir)
    send_email(to_email, subject, html_content, pdf_buffer)

    pdf_download_url = url_for(
        "download_pdf", inscricao_id=nova_inscricao.id, _external=True
    )
    return jsonify(success=True, download_url=pdf_download_url)


@app.route("/download_pdf/<int:inscricao_id>", methods=["GET"])
def download_pdf(inscricao_id):
    inscricao = Inscricao.query.get_or_404(inscricao_id)
    pdf_buffer = create_pdf(inscricao, app.config["LOGO_PATH"])
    pdf_buffer.seek(0)
    return send_file(
        pdf_buffer,
        as_attachment=True,
        download_name="Inscricao.pdf",  # Correto: 'download_name' em vez de 'attachment_filename'
        mimetype="application/pdf",
    )


def create_pdf(data, logo_path):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=18,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="LeftAlign", alignment=TA_LEFT))

    # Estrutura do documento
    story = []

    # Espaçamento acima do título
    story.append(Spacer(1, 20))  # Altere o 48 para aumentar ou diminuir o espaço

    # Adicionando título
    title = Paragraph(
        "Inscrição - Prêmio CNA Brasil Artesanal 2024 - Cerveja", styles["Title"]
    )
    story.append(title)
    story.append(Spacer(1, 12))

    # Informações de inscrição formatadas
    fields = [
        ("Nome Completo", data.nome_completo),
        ("CPF/CNPJ", data.cpf),
        ("Email", data.email),
        ("Telefone/WhatsApp", data.telefone),
        ("Número Cadastro Produtor", data.nome_estabelecimento),
        ("Município", data.municipio),
        ("Estado", data.estado),
        ("CEP", data.cep),
        ("Número Registro Inspeção", data.registro_estabelecimento_mapa),
        ("Nome Fantasia", data.nome_produto),
        ("Lote Data Envase", data.lote),
        ("Categoria Inscrita", data.categoria_inscrita),
        ("Capacidade Produtiva kg", str(data.volume_producao_anual)),
        ("Quantidade Gramas Embalagem", str(data.quantidade_ml_amostral)),
        ("História Produtor", data.historia_producao),
        ("Origem Conhecimento", data.origem_conhecimento),
        ("Outro Origem Conhecimento", data.outro_origem_conhecimento),
        ("Data Hora Inscrição", data.data_hora_inscricao.strftime("%Y-%m-%d %H:%M:%S")),
        ("Aceitou Termos", "Sim" if data.aceitou_termos else "Não"),
    ]

    for label, value in fields:
        paragraph = Paragraph(f"<b>{label}:</b> {value}", styles["BodyText"])
        story.append(paragraph)
        story.append(Spacer(1, 6))

    # Adicionando logo no cabeçalho
    def header_footer(canvas, doc):
        canvas.saveState()
        width, height = A4
        canvas.drawImage(
            logo_path,
            inch,
            height - inch - 0.5,
            width=1 * inch,
            height=0.5 * inch,
            mask="auto",
        )
        # Adicionando data, hora e paginação no rodapé
        canvas.drawString(
            inch,
            0.75 * inch,
            f"Data/Hora: {data.data_hora_inscricao.strftime('%Y-%m-%d %H:%M:%S')} - Página: {doc.page}",
        )
        canvas.restoreState()

    doc.build(story, onFirstPage=header_footer, onLaterPages=header_footer)

    buffer.seek(0)
    return buffer


def send_email(to_email, subject, html_content, pdf_file=None):
    with app.app_context():
        msg = Message(subject, recipients=[to_email], html=html_content)
        if pdf_file:
            pdf_file.seek(0)
            msg.attach("Inscricao.pdf", "application/pdf", pdf_file.read())
        mail.send(msg)


@app.route("/inscricoes", methods=["GET"])
def listar_inscricoes():
    inscricoes = Inscricao.query.all()
    return render_template("listar_inscricoes.html", inscricoes=inscricoes)


@app.route("/delete/<int:id>", methods=["POST"])
def delete_inscricao(id):
    # Encontrar a inscrição pelo ID
    inscricao = Inscricao.query.get(id)
    if not inscricao:
        return jsonify({"error": "Inscrição não encontrada"}), 404

    # Deletar a inscrição e confirmar as mudanças no banco de dados
    db.session.delete(inscricao)
    db.session.commit()

    # Enviar resposta de sucesso
    return jsonify({"success": "Inscrição deletada com sucesso"}), 200


# Rota para baixar os dados em Excel
@app.route("/download_excel", methods=["GET"])
def download_excel():
    query_sets = Inscricao.query.all()
    data = [inscricao.to_dict() for inscricao in query_sets]

    output = io.BytesIO()
    sheet = p.get_sheet(records=data)
    sheet.save_to_memory("xlsx", output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="Inscricoes.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)
