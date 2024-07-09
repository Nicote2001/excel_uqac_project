import total

#class def chaque montant pour chaque compte
class bilan_operations:
    def __init__(self,account,name,type, st_dominique_amount, sainte_famille_amount):
        self.account = account
        self.name = name
        self.type = type
        self.st_dominique_amount = st_dominique_amount
        self.sainte_famille_amount = sainte_famille_amount
        self.saint_gerard_amount = 0
        self.sainte_therese_amount = 0


