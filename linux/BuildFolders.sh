#!/bin/bash


cd ddf

for f in *.mp3
do
    echo "$f"
    # Anfangsbuchstaben groß
    fg=$"$(echo -e "${f}" | sed -r 's/(\<[a-zA-Z])/\U\1/g' )"
    #Spaces weg
    fg=$"$(echo -e "${fg}" | tr -d '[:space:]')"
    # Split
    A="$(cut -d'-' -f2 <<<$fg)"
    A="$(cut -c 1-3 <<<$A)"
    B="$(cut -d'-' -f3- <<<$fg)"
    #Mp3 weg
    B="${B%.Mp3}"
    echo "$A"
    echo "$B"
    
    #new tree
    Folder="$A"-"$B"
    echo "$Folder"
    mkdir -p TT
    mkdir -p TT/$Folder
    
    cp "$f" TT/"$Folder"
done
