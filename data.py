import camelot

tablas = camelot.read_pdf("test.pdf",pages='all',flavor='lattice')

tablas.export('test121.xlsx', f='excel')