#!/bin/sh

python ensymble_python2.5-0.28.py py2sis --caps="LocalServices+NetworkServices+ReadUserData+WriteUserData+UserEnvironment" --appname="calsync" --icon="icon.svg" --runinstall --verbose src calsync.sis
