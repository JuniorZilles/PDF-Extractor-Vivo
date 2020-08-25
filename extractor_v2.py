import tabula

def main():
    file = 'PDF/agosto.pdf'
    tables = tabula.read_pdf(file, pages = [2,3,4], multiple_tables = True)
    df = [x if x['VEJA OS NÚMEROS VIVO E PLANOS QUE COMPÕEM A SUA CONTA'] for x in tables]
    print(df)
main()