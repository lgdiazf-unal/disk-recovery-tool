#! /bin/bash

CARPETA=$1
BIN_PYTHON=$2
SCRIPT_PYTHON=$3
CARPETA_DESTINO=$4

mkdir -p $CARPETA_DESTINO

find "$CARPETA" -type f -name "*.xlsx" | while IFS= read -r archivo; do
    
    NOMBRE_ARCHIVO=$(basename "$archivo")

    echo "$BIN_PYTHON $SCRIPT_PYTHON $archivo $CARPETA_DESTINO/$NOMBRE_ARCHIVO"

done
