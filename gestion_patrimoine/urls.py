#importer les bibliothèque nécessaire
from django.contrib import admin
from django.urls import path
from patrimoine.views import ajout_pv_recensement, home, ajout_etat_apreciatif, ajout_region, ajout_services, afficher_regions, afficher_services, ajout_inventaire, ajout_ordre_entree, ajout_attestation_prise_en_charge, ajout_ordre_sortie, PVRecensement_table, EtatApreciatif_table, Inventaire_table, OrdreEntree_table, OrdreSortie_table, AttestationPriseEnCharge_table, recapitulatif_inventaire, export_inventaire, export_pvrecensement, export_attestation_prise_en_charge, export_etat_apreciatif, export_ordre_entree, export_ordre_sortie, Deconnexion, ajout_pv_evaluation, pv_evaluation, export_pvevaluation

#gestion des urls, on ajoute tout les urls ici
urlpatterns = [
    path('admin/', admin.site.urls),
    path('', home, name="home"),
    path('logout/', Deconnexion, name="logout"),
    #lien afficher dans le navigateur/, fonction correspondant dans views/, nom du lien
    path('ajout_pv_recensement/<int:service_id>', ajout_pv_recensement, name="ajout_pv_recensement"),
    path('ajout_pv_evaluation/<int:service_id>', ajout_pv_evaluation, name="ajout_pv_evaluation"),
    path('ajout_etat_apreciatif/<int:service_id>', ajout_etat_apreciatif, name="ajout_etat_apreciatif"),
    path('ajout_inventaire/<int:service_id>', ajout_inventaire, name="ajout_inventaire"),
    path('ajout_ordre_entree/<int:service_id>', ajout_ordre_entree, name="ajout_ordre_entree"),
    path('ajout_ordre_sortie/<int:service_id>', ajout_ordre_sortie, name="ajout_ordre_sortie"),
    path('ajout_attestation_prise_en_charge/<int:service_id>', ajout_attestation_prise_en_charge, name="ajout_attestation_prise_en_charge"),
    path('pv_recensement/<int:service_id>', PVRecensement_table, name="pv_recensement"),
    path('pv_evaluation/<int:service_id>', pv_evaluation, name="pv_evaluation"),
    path('etat_appreciatif/<int:service_id>', EtatApreciatif_table, name="etat_appreciatif"),
    path('inventaire/<int:service_id>', Inventaire_table, name="inventaire"),
    path('ordre_entre/<int:service_id>', OrdreEntree_table, name="ordre_entree"),
    path('ordre_sortie/<int:service_id>', OrdreSortie_table, name="ordre_sortie"),
    path('attestation_prise_en_charge/<int:service_id>', AttestationPriseEnCharge_table, name="attestation"),
    path('recapitulatif/<int:service_id>', recapitulatif_inventaire, name='recapitulatif_inventaire'),
    path('export_inventaire/<int:service_id>', export_inventaire, name='export_inventaire'),
    path('export_pvrecensement/<int:service_id>', export_pvrecensement, name='export_pvrecensement'),
    path('export_pvevaluation/<int:service_id>', export_pvevaluation, name='export_pvevaluation'),
    path('export_attestation_prise_en_charge/<int:service_id>', export_attestation_prise_en_charge, name='export_attestation_prise_en_charge'),
    path('export_etat_apreciatif/<int:service_id>', export_etat_apreciatif, name='export_etat_apreciatif'),
    path('export_ordre_entree/<int:service_id>', export_ordre_entree, name='export_ordre_entree'),
    path('export_ordre_sortie/<int:service_id>', export_ordre_sortie, name='export_ordre_sortie'),
    path('ajout_region/', ajout_region, name='ajout_region'),
    path('ajout_services/<int:region_id>/', ajout_services, name='ajout_services'),
    path('afficher_regions/', afficher_regions, name='afficher_regions'),
    path('afficher_services/<int:region_id>/', afficher_services, name='afficher_services'),
]

