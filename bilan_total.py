from total import total

#class def
class bilan_totals:
    def __init__(self, list_bilan):
        self.list_bilan = list_bilan
        self.paroissien_total = total(0,0)
        self.autre_total = total(0,0)
        self.revenus_total = total(0,0)
        self.pastorale_total = total(0,0)
        self.bureau_total = total(0,0)
        self.batisse_total = total(0,0)
        self.depenses_total = total(0,0)

    def set_totaux(self):
        for x in range(0,len(self.list_bilan)):
            if(x < 11):
                self.paroissien_total.st_dominique += self.list_bilan[x].st_dominique_amount
                self.paroissien_total.st_famille += self.list_bilan[x].sainte_famille_amount
            elif(x < 16):
                self.autre_total.st_dominique += self.list_bilan[x].st_dominique_amount
                self.autre_total.st_famille += self.list_bilan[x].sainte_famille_amount
            elif(x < 30):
                self.pastorale_total.st_dominique += self.list_bilan[x].st_dominique_amount
                self.pastorale_total.st_famille += self.list_bilan[x].sainte_famille_amount
            elif(x < 38):
                self.bureau_total.st_dominique += self.list_bilan[x].st_dominique_amount
                self.bureau_total.st_famille += self.list_bilan[x].sainte_famille_amount
            elif(x < 45):
                self.depenses_total.st_dominique += self.list_bilan[x].st_dominique_amount
                self.depenses_total.st_famille += self.list_bilan[x].sainte_famille_amount

        self.revenus_total.st_dominique = self.paroissien_total.st_dominique + self.autre_total.st_dominique
        self.revenus_total.st_famille = self.paroissien_total.st_famille + self.autre_total.st_famille

        self.depenses_total.st_dominique = self.pastorale_total.st_dominique + self.bureau_total.st_dominique + self.batisse_total.st_dominique
        self.depenses_total.st_famille = self.pastorale_total.st_famille + self.bureau_total.st_famille + self.batisse_total.st_famille
            


