from flask import Flask, render_template, request
import pickle

app = Flask(__name__,template_folder = r"Jayanthi061993/ml-deploy-documentclassification/") 
                                                                               
    
clf = pickle.load(open('model.pkl','rb'))
loaded_vec = pickle.load(open("fitted_vectorizer.pkl", "rb"))

@app.route('/')

def symptom():
    return render_template('category_pred.html')

@app.route('/result',methods = ['POST','GET'])
def result():
    if request.method == 'POST':
        result = request.form['Data']
        result_pred = clf.predict(loaded_vec.transform([result]))
        return render_template('category_result.html',result = result_pred)

if __name__ == '__main__':
    app.run()
