from total import total

#class def sert a calculer les grands totaux
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
        self.surplus_annee = total(0,0)
        self.surplus_accumule = total(0,0)

    def set_totaux(self):
        for x in range(0,len(self.list_bilan)):
            if(x < 11):
                self.paroissien_total.st_dominique += self.list_bilan[x].st_dominique_amount
                self.paroissien_total.st_famille += self.list_bilan[x].sainte_famille_amount
                self.paroissien_total.st_gerard += self.list_bilan[x].saint_gerard_amount
                self.paroissien_total.st_therese += self.list_bilan[x].sainte_therese_amount
            elif(x < 16):
                self.autre_total.st_dominique += self.list_bilan[x].st_dominique_amount
                self.autre_total.st_famille += self.list_bilan[x].sainte_famille_amount
                self.autre_total.st_gerard += self.list_bilan[x].saint_gerard_amount
                self.autre_total.st_therese += self.list_bilan[x].sainte_therese_amount
            elif(x < 30):
                self.pastorale_total.st_dominique += self.list_bilan[x].st_dominique_amount
                self.pastorale_total.st_famille += self.list_bilan[x].sainte_famille_amount
                self.pastorale_total.st_gerard += self.list_bilan[x].saint_gerard_amount
                self.pastorale_total.st_therese += self.list_bilan[x].sainte_therese_amount
            elif(x < 37):
                self.bureau_total.st_dominique += self.list_bilan[x].st_dominique_amount
                self.bureau_total.st_famille += self.list_bilan[x].sainte_famille_amount
                self.bureau_total.st_gerard += self.list_bilan[x].saint_gerard_amount
                self.bureau_total.st_therese += self.list_bilan[x].sainte_therese_amount
            elif(x < 45):
                self.batisse_total.st_dominique += self.list_bilan[x].st_dominique_amount
                self.batisse_total.st_famille += self.list_bilan[x].sainte_famille_amount
                self.batisse_total.st_gerard += self.list_bilan[x].saint_gerard_amount
                self.batisse_total.st_therese += self.list_bilan[x].sainte_therese_amount

        self.revenus_total.st_dominique = self.paroissien_total.st_dominique + self.autre_total.st_dominique
        self.revenus_total.st_famille = self.paroissien_total.st_famille + self.autre_total.st_famille
        self.revenus_total.st_gerard = self.paroissien_total.st_gerard + self.autre_total.st_gerard
        self.revenus_total.st_therese = self.paroissien_total.st_therese + self.autre_total.st_therese

        self.depenses_total.st_dominique = self.pastorale_total.st_dominique + self.bureau_total.st_dominique + self.batisse_total.st_dominique
        self.depenses_total.st_famille = self.pastorale_total.st_famille + self.bureau_total.st_famille + self.batisse_total.st_famille
        self.depenses_total.st_gerard = self.pastorale_total.st_gerard + self.bureau_total.st_gerard + self.batisse_total.st_gerard
        self.depenses_total.st_therese = self.pastorale_total.st_therese + self.bureau_total.st_therese + self.batisse_total.st_therese

        self.surplus_accumule.st_dominique = 0
        self.surplus_accumule.st_famille = 0
        self.surplus_accumule.st_gerard = 0
        self.surplus_accumule.st_therese = 0

        self.surplus_annee.st_dominique = self.revenus_total.st_dominique -self.depenses_total.st_dominique
        self.surplus_annee.st_famille = self.revenus_total.st_famille -self.depenses_total.st_famille
        self.surplus_annee.st_gerard = self.revenus_total.st_gerard -self.depenses_total.st_gerard
        self.surplus_annee.st_therese = self.revenus_total.st_therese -self.depenses_total.st_therese
        
            


