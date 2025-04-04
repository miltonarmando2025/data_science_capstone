import os
import nbformat
from nbconvert.preprocessors import ExecutePreprocessor

def execute_notebook(notebook_path):
    # Load the notebook
    with open(notebook_path) as f:
        nb = nbformat.read(f, as_version=4)
    
    # Execute the notebook
    ep = ExecutePreprocessor(timeout=600, kernel_name='python3')
    ep.preprocess(nb, {'metadata': {'path': os.path.dirname(notebook_path)}})
    

notebook_path = [
                "notebook\Data_Wrangling.ipynb",
                "notebook\EDA_Visualization.ipynb",
                "notebook\Predictive_Analysis.ipynb"
                ]


os.system("python src\generate_presentation.py")

if __name__ == "__main__":
    # Execute the notebook
    for n in notebook_path:
        execute_notebook(n)
    os.system("python src\generate_presentation.py")