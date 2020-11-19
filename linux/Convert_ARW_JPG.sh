for i in *.ARW
do 
	ufraw-batch --wb=camera --exposure=auto --lensfun=none --out-type=jpeg "$i" 
	#NAME=`echo "$i" | sed "s/.ARW//g"`
	#convert $i $NAME.jpg
done 
