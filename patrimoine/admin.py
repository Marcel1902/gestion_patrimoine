
from django.contrib import admin
from  django.utils.translation import gettext_lazy
from .models import Inventaire, PVRecensement, OrdreEntree, OrdreSortie, EtatApreciatif, AttestationPriseEnCharge, Region, Service
# Enregistrer notre table dans l'espace admin pour pouvoir les gérer dans l'espace admin
class InventaireAdmin(admin.ModelAdmin):
    list_display = ('Designation_materiels', 'Especes_des_unites', 'Prix_de_l_unite', 'Quantite_existant_1er_janvier', 'Quantite_entree_pendant_l_annee', 'Quantite_sortie_pendant_l_annee', 'Quantite_reste_31_decembre', 'Decompte')
    list_filter = ('Annee_exercice', 'Nomenclature')

class EtatApreciatifAdmin(admin.ModelAdmin):
    list_display = ('Annee_exercice', 'Nomenclature', 'Numero_du_piece_justificative', 'Date_du_piece_justificative', 'Designations_sommaire_des_operations', 'Charge', 'Decharge')
    list_filter = ('Annee_exercice', 'Nomenclature')

class PVRecensementAdmin(admin.ModelAdmin):
    list_display = ('Service', 'Annee_exercice', 'Nomenclature', 'Designation_materiels', 'Especes_unites', 'Prix_unites', 'Quantites_d_apres_ecriture', 'Quantites_par_recensement', 'Quantites_execedent_par_article', 'Quantites_deficient_par_article', 'valeurs_excedents_par_article', 'valeurs_excedents_par_nomenclature', 'valeurs_deficits_par_article', 'valeurs_des_deficits_par_nomenclature', 'valeurs_des_existants', 'Observations')
    list_filter = ('Annee_exercice', 'Nomenclature')

class OrdreEntreeAdmin(admin.ModelAdmin):
    list_display = ('Annee_exercice', 'Numero_folio_du_grandlivre', 'Nomenclature', 'Designation_des_matieres_et_objets', 'Especes_des_unites', 'Quantites', 'Prix_unite', 'Valeurs_partielles', 'Valeurs_par_numero_nomenclature', 'Numero_piece_justificative_sortie_correspondante')
    list_filter = ('Annee_exercice', 'Nomenclature')

class OrdreSortieAdmin(admin.ModelAdmin):
    list_display = ('Annee_exercice', 'Numero_folio_du_grandlivre', 'Nomenclature', 'Designation_des_matieres_et_objets', 'Especes_des_unites', 'Quantites', 'Prix_unite', 'Valeurs_partielles', 'Valeurs_par_numero_nomenclature', 'Numero_piece_justificative_sortie_correspondante')
    list_filter = ('Annee_exercice', 'Nomenclature')

class AttestationPriseEnChargeAdmin(admin.ModelAdmin):
    list_display = ('Service', 'Annee_exercice', 'Nomenclature', 'Designation_des_matieres_et_objets', 'Especes_des_unites', 'Quantite', 'Prix_unite', 'Montant', 'Observations')
    list_filter = ('Annee_exercice', 'Nomenclature')

admin.site.register(Inventaire, InventaireAdmin)
admin.site.register(PVRecensement, PVRecensementAdmin)
admin.site.register(OrdreEntree, OrdreEntreeAdmin)
admin.site.register(OrdreSortie, OrdreSortieAdmin)
admin.site.register(EtatApreciatif, EtatApreciatifAdmin)
admin.site.register(AttestationPriseEnCharge, AttestationPriseEnChargeAdmin)
admin.site.register(Region)
admin.site.register(Service)

admin.site.site_title = gettext_lazy("GESTION PATRIMOINE")
admin.site.site_header = gettext_lazy("GESTION PATRIMOINE")
admin.site.index_title = gettext_lazy("GESTION PATRIMOINE")