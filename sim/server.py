print("Remember to run pip install -r requirements.txt first")
import sys
import json
from flask import Flask
from flask_cors import CORS
from flask import request
import spacy_universal_sentence_encoder

print("If this is the first time running this it might take a while")

# https://spacy.io/universe/project/spacy-universal-sentence-encoder
nlp = spacy_universal_sentence_encoder.load_model('en_use_lg')


app = Flask(__name__)
CORS(app)
cors = CORS(app, resource={
    r"/*":{
        "origins":"*"
    }
})

@app.route('/paragraph-similarity',methods = ['POST'])
def paragraphs_endpoint():
    data = json.loads(request.data)
    print(data)
    paragraphs = data['paragraphs']
    alttext = data['alttext']

    alttext = nlp(alttext)
    sims = [(alttext.similarity(nlp(p)), p) for p in paragraphs]
    sims.sort(key=lambda x:-x[0])
    
    print(sims)
    print(paragraphs)
    print(alttext)

    r = { 'sims': sims }
    
    response = app.response_class(
            response=json.dumps(r),
            status=200,
            mimetype='application/json'
    )
    return response


if __name__ == '__main__':
    print(sys.argv)
    port = 5001 if len(sys.argv) < 2 else int(sys.argv[1].strip())
    app.run(host='localhost', port=port)