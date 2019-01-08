echo "Starting project initialization. This sould only be run once per machine!"
echo "Creating a local python environment in .venv and activating it"
python3 -m venv .venv
. activate.sh

echo "Installing requirements"
pip install --upgrade pip
pip install -r requirements.txt

echo "You should now have a local python3 version:"
python --version
which python

echo "Your environment should contain numpy:"
pip list --format=columns | grep numpy
