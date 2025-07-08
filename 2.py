import segno

# Create a Data Matrix
dm = segno.make("X1824", micro=True)  # micro=False for standard Data Matrix
dm.save("datamatrix.png", scale=5)  # Increase `scale` to make it larger