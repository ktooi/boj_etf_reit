import unittest
from boj_etf_reit.reader import workbook


class BojEtfReitTestHelper(object):

    mul = 10000000000

    @classmethod
    def etf_jreit(cls, trade_date, etf, jreit, l=None):
        r = {}
        r['trade_date'] = trade_date
        r['etf'] = {'other': etf[0] * cls.mul, 'support': etf[1] * cls.mul}
        r['reit'] = jreit * cls.mul
        if l is not None:
            l.append(r)
        return r


class TestEtfReitXlsDriver(unittest.TestCase):

    def test_read(self):
        etfreit = workbook.EtfReitXlsDriver('boj_etf_reit\\tests\\data\\etfreit04.xlsx')
        act = etfreit.read()

        exp = []
        BojEtfReitTestHelper.etf_jreit("2020-04-01", [1202, 12], 20, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-02", [1202, 12], 20, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-03", [0, 12], 20, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-04", [0, 0], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-05", [0, 0], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-06", [0, 12], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-07", [0, 12], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-08", [0, 12], 20, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-09", [1202, 12], 20, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-10", [1202, 12], 20, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-11", [0, 0], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-12", [0, 0], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-13", [0, 12], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-14", [0, 12], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-15", [1202, 12], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-16", [1202, 12], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-17", [0, 12], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-18", [0, 0], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-19", [0, 0], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-20", [0, 12], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-21", [1202, 12], 20, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-22", [1202, 12], 20, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-23", [0, 12], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-24", [1202, 12], 20, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-25", [0, 0], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-26", [0, 0], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-27", [0, 12], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-28", [1202, 12], 20, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-29", [0, 0], 0, exp)
        BojEtfReitTestHelper.etf_jreit("2020-04-30", [0, 12], 0, exp)

        self.assertEqual(act, exp, "Read and out")

    
if __name__ == '__main__':
    unittest.main()