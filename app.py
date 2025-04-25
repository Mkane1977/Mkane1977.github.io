from flask import Flask, render_template, request
import spacy
import re

app = Flask(__name__)

# 1) Load spaCy model once
nlp = spacy.load("en_core_web_sm")

# 2) Entity labels & colors
LABELS = ["PERSON","ORG","GPE","DATE","TIME","PRODUCT","MONEY","QUANTITY",
          "PHONE","EMAIL","URL","PERCENT"]
COLORS = {
  "PERSON":"#C8E6C9","ORG":"#A2C8EC","GPE":"#FFCDD2","DATE":"#FFF9C4",
  "TIME":"#F8BBD0","PRODUCT":"#B2EBF2","MONEY":"#D1C4E9","QUANTITY":"#C5CAE9",
  "PHONE":"#E0F7FA","EMAIL":"#E1BEE7","URL":"#DCEDC8","PERCENT":"#FFCCBC"
}

def preprocess(text):
    # insert periods if none (helps sentence detection; optional)
    if not any(p in text for p in ".!?"):
        tokens = text.split()
        out=[]
        for i,t in enumerate(tokens):
            out.append(t)
            if i<len(tokens)-1 and re.match(r"^\$?\d+(\.\d+)?$",t) and tokens[i+1][0].isupper():
                out.append(".")
        text=" ".join(out)
    return text

@app.route("/", methods=["GET","POST"])
def index():
    return render_template("index.html", labels=LABELS, colors=COLORS)

@app.route("/ner", methods=["POST"])
def ner():
    text = request.form["text"]
    entities = request.form.getlist("entities")  # list of selected labels
    text = preprocess(text)
    doc = nlp(text)

    # Build highlighted HTML
    result = []
    last = 0
    for ent in doc.ents:
        if ent.label_ in entities:
            # add plain text before ent
            result.append(text[last:ent.start_char])
            color = COLORS.get(ent.label_, "#FFFF00")
            # wrap the entity + label
            span = f'<span class="entity" style="background:{color}">' \
                   f'{ent.text} [{ent.label_}]</span>'
            result.append(span)
            last = ent.end_char
    result.append(text[last:])
    highlighted = "".join(result)

    return render_template("result.html",
                           labels=LABELS,
                           colors=COLORS,
                           highlighted=highlighted,
                           request=request)

if __name__ == "__main__":
    app.run(debug=True)
