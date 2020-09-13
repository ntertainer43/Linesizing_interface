
import xlwings as xw


class Sizing(object):
    """docstring for Sizing."""

    def __init__(self, flpath):
        super(Sizing, self).__init__()
        self.flpath = 'C:/Users/capta/Desktop/Linesizing/Linesizing/HMB.xlsx'
        self.wb = xw.Book(self.flpath)
        self.sh = self.wb.sheets['Property Calc_Min. Flow']
        self.streamlist = []
        self.prop = self.sh.range('C4:C13').value
        self.unit = self.sh.range('D4:D13').value
        print(self.unit)
        self.propvalue =[]
        self.parsestream()

    def parsestream(self):
        j = 5
        while  self.sh.range(2,j).value != None:
            self.readstream(j)
            self.streamlist.append(self.sh.range(2,j).value)
            j += 6


    def readstream(self, j):
        a =[]
        for i in range (3,13):
            a.append(self.sh.range(i,j).value)
            #print(a)
        self.propvalue.append(a)
        print(self.propvalue)





def main():
    pipe = Sizing('C:/Users/capta/Desktop/Linesizing/Linesizing/HMB.xlsx')

if __name__ == '__main__':
    main()
