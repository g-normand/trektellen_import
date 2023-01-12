
ENV_PATH := $(CURDIR)/python_env
PYTHON := $(ENV_PATH)/bin/python3.9
VENV = source $(ENV_PATH)/bin/activate

# For shell to bash to be able to use source.
SHELL = /bin/bash

# Virtualenv file containing python libraries.
virtualenv:
	virtualenv -p /usr/bin/python3.9 $(ENV_PATH)

# Install python requirements.
install: virtualenv
	$(VENV) && cd $(APP_PATH) && pip3 install -r $(CURDIR)/requirements.txt;

trektellen: virtualenv
	$(VENV) && python3 trektellen.py
