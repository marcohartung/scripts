#!/bin/bash
for i in ./*.JPG
do 
	echo "$i"
	convert $i -quality 50% -resize 50% ./out/$i.jpg
done 
