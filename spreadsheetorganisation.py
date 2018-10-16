import openpyxl

def main():
    opt = parse_cmdline()
    fix_file(opt.filein, opt.fileout)

def fix_file(filein, fileout):
    wb = openpyxl.load_workbook(filein)
    for ws in wb:
        thisdict = {}
        dupedict = {}
        for row in ws:
            for cell in row:
                if not cell.hyperlink:
                    continue
                # if cell text is same as other cell text, print cell
                if cell.value in thisdict:
                    print("Dupe " + cell.value + " HL " + cell.hyperlink.target)
                    if cell.value in dupedict:
                        dupedict[cell.value] = [thisdict[thisdict[cell.value]]]
                    dupedict[cell.value] = dupedict[cell.value].append(cell.hyperlink.target)
                    # above three lines possibly incorrect
                    # next step is to list all dupe hyperlinks
                else:
                    thisdict[cell.value] = cell.hyperlink.target
                # also remove file from file path in dupe cell

    wb.save(fileout)


def parse_cmdline():
    from argparse import ArgumentParser
    parser = ArgumentParser()
    parser.add_argument('filein')
    parser.add_argument('fileout')
    opt = parser.parse_args()
    return opt

if __name__ == '__main__':
    main()
