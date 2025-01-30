Typing game made with excel VBA :

rules : type the word written on the apple before it falls on the ground 

How to : 
1. download the folder 
2. open project 
3. accept macro on excel 
4. modify macro named Module1 
5. modify the line .Picture = LoadPicture("C:\Users\C213\Desktop\vba\apple.jpg") 
	replace the path with the path of the picture apple in the folder 
5. Start the game with the button 'Jouer'

-------------------------------------------------------------------------------------------------

-Pour jouer : appuyer sur le bouton jouer de la feuille 1 

-Dans le module1, dans la macro attente: 
	la ligne  .Picture = LoadPicture("C:\Users\C213\Desktop\vba\apple.jpg") 
	LoadPicture() doit contenir ENTRE GUILLEMETS le bon CHEMIN du fichier apple.jpg qui se trouve dans le dossier
