// Racuna se y vrijednost za broj na x osi
// Poziva se methoda SetVal koja je definirana u host aplikaciji koja prima 2 argumenta
for (i = 0; i < 500; i++)
  app.SetVal(i, Math.cos(i * Math.PI / 90) * 200 + 220);

// Renderiranje
app.Invalidate();

// Dohvacanje vrijednosti za vrijednost na osi x
// Poziva se funkcija GetVal iz host-a koja prima vrijednost za os x i vraca vrijednost na osi y
x = 20;
y = app.GetVal(x);

// Poziva se MessageBoxA koja je definirana u user32.dll
app.MessageBoxA(0, '(' + x + ', ' + y + ')', 'Vrijednost (X, Y)', 0);