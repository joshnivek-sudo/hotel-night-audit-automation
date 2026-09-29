"""
Testes da automação do Estatístico + Folha de Rosto.

Rodar (na pasta do projeto):
    python -m unittest discover -s tests -v
"""
import datetime
import importlib.machinery
import importlib.util
import os
import pathlib
import shutil
import tempfile
import unittest

import openpyxl

RAIZ = pathlib.Path(__file__).resolve().parents[1]
EXEMPLOS = RAIZ / "exemplos" / "estatistico"
DATA = datetime.datetime(2026, 9, 28)


def carregar_modulo():
    # o script é um .bat que também é Python válido: carrega como módulo
    caminho = str(RAIZ / "estatistico" / "Preencher Estatistico.bat")
    loader = importlib.machinery.SourceFileLoader("preencher_estatistico", caminho)
    spec = importlib.util.spec_from_loader("preencher_estatistico", loader)
    modulo = importlib.util.module_from_spec(spec)
    loader.exec_module(modulo)
    return modulo


est = carregar_modulo()


class TestUtilidades(unittest.TestCase):

    def test_limpar_tira_acento_e_pontuacao(self):
        self.assertEqual(est.limpar("Hóspedes Adultos/Criança1"), "HOSPEDES ADULTOS CRIANCA1")
        self.assertEqual(est.limpar("UH's Bloqueadas"), "UH'S BLOQUEADAS")

    def test_num_entende_formato_brasileiro(self):
        self.assertEqual(est.num("1.234,56"), 1234.56)
        self.assertEqual(est.num("89,90"), 89.90)
        self.assertEqual(est.num(42), 42.0)
        self.assertIsNone(est.num("abc"))

    def test_nome_com_data(self):
        dia = datetime.datetime(2026, 9, 6)
        self.assertEqual(os.path.basename(est.nome_com_data("ESTATISTICO__05-09-2026.xlsx", dia)),
                         "ESTATISTICO__06-09-2026.xlsx")
        self.assertEqual(os.path.basename(est.nome_com_data("ESTATISTICO.xlsx", dia)),
                         "ESTATISTICO_06-09-2026.xlsx")


class TestLeituraDosRelatorios(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.rds = est.ler_rds(str(EXEMPLOS / "RDS_RELATORIO_DEMO.xlsx"))
        cls.otb = est.ler_on_the_book(str(EXEMPLOS / "on_the_book_demo.xlsx"))
        cls.cp = est.ler_conta_pendente(str(EXEMPLOS / "conta_pendente_demo.xlsx"))
        cls.uhs = est.ler_situacao_uhs(str(EXEMPLOS / "situacao_uhs_demo.xlsx"), cls.rds["data"])

    def test_rds_data(self):
        self.assertEqual(self.rds["data"], DATA)

    def test_rds_estatisticas(self):
        e = self.rds["estatisticas"]
        self.assertEqual(e[5], 40)    # UH's do hotel
        self.assertEqual(e[9], 31)    # UH's alugadas
        self.assertEqual(e[13], 52)   # adultos
        self.assertEqual(e[14], 4)    # crianças (3 + 1)
        self.assertEqual(e[11], 1)    # cortesia + permuta

    def test_rds_no_show_pega_o_numero_da_esquerda(self):
        self.assertEqual(self.rds["extra"]["no_show"], 1)

    def test_rds_ignora_recebimentos(self):
        itens = [item for _, item, _ in self.rds["itens_nao_mapeados"]]
        self.assertNotIn("DINHEIRO", itens)
        self.assertEqual(self.rds["itens_nao_mapeados"], [])

    def test_rds_receitas_batem_com_total_de_debitos(self):
        total = (sum(self.rds["receitas"].values()) + self.rds["estorno_hospedagem"]
                 + self.rds["estorno_outras"])
        self.assertAlmostEqual(total, self.rds["total_debitos"], places=2)

    def test_on_the_book_comeca_no_dia_seguinte(self):
        self.assertIn(datetime.date(2026, 9, 29), self.otb["dias"])
        self.assertNotIn(datetime.date(2026, 9, 28), self.otb["dias"])
        self.assertIsNotNone(self.otb["total_receita"])

    def test_conta_pendente(self):
        self.assertEqual(self.cp["qtd"], 7)
        self.assertGreater(self.cp["credito"], 0)

    def test_situacao_uhs_so_conta_o_pool_do_hotel(self):
        self.assertEqual(self.uhs["total_pool"], 40)   # as 8 UHs de condomínio ficam fora

    def test_situacao_uhs_foto_atual(self):
        u = self.uhs
        self.assertEqual((u["ocupadas"], u["bloqueadas"], u["vago_limpo"], u["vago_sujo"]),
                         (32, 1, 5, 2))
        self.assertEqual(sorted(u["walk_in"]), ["103", "208"])
        self.assertEqual(sorted(u["layover"]), ["305", "306"])
        self.assertEqual(u["mensalistas"], ["410"])


class TestPreenchimentoCompleto(unittest.TestCase):

    def test_preenche_estatistico_e_folha(self):
        rds = est.ler_rds(str(EXEMPLOS / "RDS_RELATORIO_DEMO.xlsx"))
        otb = est.ler_on_the_book(str(EXEMPLOS / "on_the_book_demo.xlsx"))
        cp = est.ler_conta_pendente(str(EXEMPLOS / "conta_pendente_demo.xlsx"))
        uhs = est.ler_situacao_uhs(str(EXEMPLOS / "situacao_uhs_demo.xlsx"), rds["data"])

        with tempfile.TemporaryDirectory() as tmp:
            est_path = shutil.copy(EXEMPLOS / "ESTATISTICO_HOTEL_DEMO.xlsx", tmp)
            folha_path = shutil.copy(EXEMPLOS / "FOLHA_DE_ROSTO_DEMO.xlsx", tmp)

            (out_est, out_fol, fr, on_the_book, dif,
             receitas, avisos, prev, backups) = est.preencher(est_path, folha_path,
                                                             rds, otb, cp, uhs)

            # conferência: receitas lançadas = total de débitos do RDS
            self.assertAlmostEqual(dif, 0, places=2)

            # arquivos renomeados com a data e o do dia anterior removido
            self.assertTrue(out_est.endswith("ESTATISTICO_HOTEL_DEMO_28-09-2026.xlsx"))
            self.assertTrue(out_fol.endswith("FOLHA_DE_ROSTO_DEMO_28-09-2026.xlsx"))
            self.assertFalse(os.path.exists(est_path))
            self.assertEqual(len(backups), 2)

            wb = openpyxl.load_workbook(out_est)
            lc = wb["Lanç. Estatístico"]
            self.assertEqual(wb["Estatístico"]["C4"].value, 28)
            self.assertEqual(lc["AC9"].value, 31)     # dia 28 = coluna AC
            self.assertEqual(lc["AC12"].value, 2)     # walk-ins

            folha = openpyxl.load_workbook(out_fol)["Resumo Estatistico"]
            self.assertEqual(folha["F7"].value, 28)
            self.assertAlmostEqual(folha["F8"].value, 31 / 40)          # ocupação do dia
            self.assertEqual(folha["M5"].value, 32)                     # ocupadas agora
            self.assertAlmostEqual(folha["C15"].value, 1.68)            # só adultos / ocupadas
            self.assertIsNone(folha["M13"].value)                       # check-in após 00h


if __name__ == "__main__":
    unittest.main()
