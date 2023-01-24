
ENV_PATH := $(CURDIR)/python_env
PYTHON := $(ENV_PATH)/bin/python3.9

# For shell to bash to be able to use source.
SHELL = /bin/bash

# Virtualenv file containing python libraries.
virtualenv:
	source $(ENV_PATH)/bin/activate

# Install python requirements.
install:
	virtualenv -p /usr/bin/python3.9 $(ENV_PATH) && cd $(APP_PATH) && pip3 install -r $(CURDIR)/requirements.txt;

test: virtualenv
	python3 trektellen.py ben_test_file_alaska.xls

trektellen: virtualenv
	python3 trektellen.py ben_2007.xlsx
