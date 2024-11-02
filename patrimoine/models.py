from idlelib.pyparse import trans

from django.db import models
from django.db.models import Model


# Create your models here.
class Region(models.Model):
    id = models.AutoField(primary_key=True)
    Nom = models.CharField(max_length=50)

    def __str__(self):
        return self.Nom


class Service(models.Model):
    id = models.AutoField(primary_key=True)
    Region = models.ForeignKey(Region, on_delete=models.CASCADE)
    Nom = models.CharField(max_length=50)

    def __str__(self):
        return self.Nom


class PVRecensement(models.Model):
    Budget = models.CharField(max_length=50, blank=True, null=True)
    Imputation_administrative = models.CharField(max_length=50, blank=True, null=True)
    Designation_chapitre = models.CharField(max_length=10, blank=True, null=True)
    Libelle_article = models.CharField(max_length=100, blank=True, null=True)
    Designation_magasin_ou_materiels_service = models.CharField(max_length=100, blank=True, null=True)
    Nom_et_qualite_recenseur = models.CharField(max_length=100, blank=True, null=True)
    Nom_et_qualite_depositaire_comptable = models.CharField(max_length=250, blank=True, null=True)
    Service = models.ForeignKey(Service, on_delete=models.CASCADE)
    Annee_exercice = models.CharField(max_length=50)
    Nomenclature = models.IntegerField()
    Designation_materiels = models.CharField(max_length=250)
    Especes_unites = models.CharField(max_length=50)
    Prix_unites = models.FloatField(default=0.00)
    Quantites_d_apres_ecriture = models.IntegerField(default=0)
    Quantites_par_recensement = models.IntegerField(default=0)
    Quantites_execedent_par_article = models.IntegerField(default=0)
    Quantites_deficient_par_article = models.IntegerField(default=0)
    valeurs_excedents_par_article = models.FloatField(default=0.00)
    valeurs_excedents_par_nomenclature = models.FloatField(default=0.00)
    valeurs_deficits_par_article = models.FloatField(default=0.00)
    valeurs_des_deficits_par_nomenclature = models.FloatField(default=0.00)
    valeurs_des_existants = models.FloatField(default=0.00)
    Observations = models.CharField(max_length=250)

    def __str__(self):
        return self.Designation_materiels


class Inventaire(models.Model):
    Budget = models.CharField(max_length=50, blank=True, null=True)
    imputation_administrative = models.CharField(max_length=50, blank=True, null=True)
    Designation_du_chapitre = models.CharField(max_length=100, blank=True, null=True)
    libelle_article = models.CharField(max_length=100, blank=True, null=True)
    Nom_et_qualite_du_recenseur = models.CharField(max_length=100, blank=True, null=True)
    Nom_et_qualite_du_depositaire_comptable = models.CharField(max_length=250, blank=True, null=True)
    Service = models.ForeignKey(Service, on_delete=models.CASCADE)
    Annee_exercice = models.CharField(max_length=10)
    Nomenclature = models.IntegerField()
    Numero_folio_grand_livre = models.IntegerField()
    Designation_materiels = models.CharField(max_length=250)
    Especes_des_unites = models.CharField(max_length=10)
    Prix_de_l_unite = models.FloatField(default=0.00)
    Quantite_existant_1er_janvier = models.IntegerField(default=0)
    Quantite_entree_pendant_l_annee = models.IntegerField(default=0)
    Quantite_sortie_pendant_l_annee = models.IntegerField(default=0)
    Quantite_reste_31_decembre = models.IntegerField(default=0)
    Decompte = models.FloatField(default=0.00)
    Observation = models.CharField(max_length=250)

    def __str__(self):
        return self.Designation_materiels


class EtatApreciatif(models.Model):
    Budget = models.CharField(max_length=50, blank=True, null=True)
    Materiel_en_approvisionnement_ou_en_service = models.CharField(max_length=50, blank=True, null=True)
    Designation_du_chapitre = models.CharField(max_length=100, blank=True, null=True)
    libelle_article = models.CharField(max_length=100, blank=True, null=True)
    Subdivision_du_magasin_ou_de_la_categorie_du_materiel_en_service = models.CharField(max_length=250, blank=True,
                                                                                        null=True)
    Nombre_en_toutes_lettres = models.CharField(max_length=50, blank=True, null=True)
    Nom_et_qualite_du_depositaire_comptable = models.CharField(max_length=250, blank=True, null=True)
    Service = models.ForeignKey(Service, on_delete=models.CASCADE)
    Annee_exercice = models.CharField(max_length=10)
    Nomenclature = models.IntegerField()
    Numero_du_piece_justificative = models.CharField(max_length=20)
    Date_du_piece_justificative = models.DateField()
    Designations_sommaire_des_operations = models.CharField(max_length=250)
    Charge = models.FloatField(default=0.00)
    Decharge = models.FloatField(default=0.00)

    def __str__(self):
        return self.Designations_sommaire_des_operations


class OrdreEntree(models.Model):
    Budget = models.CharField(max_length=50, blank=True, null=True)
    imputation_administrative = models.CharField(max_length=50, blank=True, null=True)
    Designation_du_chapitre = models.CharField(max_length=100, blank=True, null=True)
    libelle_article = models.CharField(max_length=100, blank=True, null=True)
    Subdivision_du_chapitre = models.CharField(max_length=250, blank=True, null=True)
    Numero_du_journal = models.CharField(max_length=10, blank=True, null=True)
    Nom_et_qualite_du_depositaire_comptable = models.CharField(max_length=250, blank=True, null=True)
    Service = models.ForeignKey(Service, on_delete=models.CASCADE)
    Annee_exercice = models.CharField(max_length=10)
    Numero_folio_du_grandlivre = models.IntegerField()
    Nomenclature = models.IntegerField()
    Designation_des_matieres_et_objets = models.CharField(max_length=250)
    Especes_des_unites = models.CharField(max_length=20)
    Quantites = models.IntegerField(default=0)
    Prix_unite = models.FloatField(default=0.00)
    Valeurs_partielles = models.FloatField(default=0.00)
    Valeurs_par_numero_nomenclature = models.FloatField(default=0.00)
    Numero_piece_justificative_sortie_correspondante = models.CharField(max_length=20)

    def __str__(self):
        return self.Designation_des_matieres_et_objets


class AttestationPriseEnCharge(models.Model):
    Service = models.ForeignKey(Service, on_delete=models.CASCADE)
    Annee_exercice = models.CharField(max_length=50)
    Nomenclature = models.IntegerField()
    Designation_des_matieres_et_objets = models.CharField(max_length=250)
    Especes_des_unites = models.CharField(max_length=20)
    Quantite = models.IntegerField(default=0)
    Prix_unite = models.FloatField(default=0.00)
    Montant = models.FloatField(default=0.00)
    Observations = models.CharField(max_length=250)

    def __str__(self):
        return self.Designation_des_matieres_et_objets


class OrdreSortie(models.Model):
    Budget = models.CharField(max_length=50, blank=True, null=True)
    imputation_administrative = models.CharField(max_length=50, blank=True, null=True)
    Designation_du_chapitre = models.CharField(max_length=100, blank=True, null=True)
    libelle_article = models.CharField(max_length=100, blank=True, null=True)
    Subdivision_du_chapitre = models.CharField(max_length=250, blank=True, null=True)
    Numero_du_journal = models.CharField(max_length=10, blank=True, null=True)
    Nom_et_qualite_du_depositaire_comptable = models.CharField(max_length=250, blank=True, null=True)
    Service = models.ForeignKey(Service, on_delete=models.CASCADE)
    Annee_exercice = models.CharField(max_length=10)
    Numero_folio_du_grandlivre = models.IntegerField()
    Nomenclature = models.IntegerField()
    Designation_des_matieres_et_objets = models.CharField(max_length=250)
    Especes_des_unites = models.CharField(max_length=20)
    Quantites = models.IntegerField(default=0)
    Prix_unite = models.FloatField(default=0.00)
    Valeurs_partielles = models.FloatField(default=0.00)
    Valeurs_par_numero_nomenclature = models.FloatField(default=0.00)
    Numero_piece_justificative_sortie_correspondante = models.CharField(max_length=20)

    def __str__(self):
        return self.Designation_des_matieres_et_objets


class PvEvaluation(models.Model):
    Budget = models.CharField(max_length=50, blank=True, null=True)
    imputation_administrative = models.CharField(max_length=50, blank=True, null=True)
    Designation_du_chapitre = models.CharField(max_length=100, blank=True, null=True)
    libelle_article = models.CharField(max_length=100, blank=True, null=True)
    President_commission_recensement = models.CharField(max_length=100, blank=True, null=True)
    Membres = models.CharField(max_length=250, blank=True, null=True)
    Nom_et_qualite_du_depositaire_comptable = models.CharField(max_length=250, blank=True, null=True)
    Service = models.ForeignKey(Service, on_delete=models.CASCADE)
    Annee_exercice = models.CharField(max_length=10)
    Nomenclature = models.IntegerField()
    Numero_ordre = models.CharField(max_length=10)
    Designations_des_articles = models.CharField(max_length=250)
    Especes_des_unites = models.CharField(max_length=50)
    Prix_unitaire = models.FloatField(default=0.00)
    Quantite = models.IntegerField(default=0)
    Montant = models.FloatField(default=0.00)
    Observations = models.CharField(max_length=250)


