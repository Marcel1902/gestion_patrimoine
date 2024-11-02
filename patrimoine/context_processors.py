from .models import PVRecensement, EtatApreciatif, Inventaire, OrdreEntree, AttestationPriseEnCharge, OrdreSortie

def filters_context(request):
    # Récupérer les valeurs distinctes pour les filtres à partir de plusieurs tables
    nomenclature_list = set()
    annee_exercice_list = set()

    # Récupérer les valeurs distinctes de chaque modèle et les ajouter à la liste
    nomenclature_list.update(PVRecensement.objects.values_list('Nomenclature', flat=True).distinct())
    nomenclature_list.update(EtatApreciatif.objects.values_list('Nomenclature', flat=True).distinct())
    nomenclature_list.update(Inventaire.objects.values_list('Nomenclature', flat=True).distinct())
    nomenclature_list.update(OrdreEntree.objects.values_list('Nomenclature', flat=True).distinct())
    nomenclature_list.update(AttestationPriseEnCharge.objects.values_list('Nomenclature', flat=True).distinct())
    nomenclature_list.update(OrdreSortie.objects.values_list('Nomenclature', flat=True).distinct())

    annee_exercice_list.update(PVRecensement.objects.values_list('Annee_exercice', flat=True).distinct())
    annee_exercice_list.update(EtatApreciatif.objects.values_list('Annee_exercice', flat=True).distinct())
    annee_exercice_list.update(Inventaire.objects.values_list('Annee_exercice', flat=True).distinct())
    annee_exercice_list.update(OrdreEntree.objects.values_list('Annee_exercice', flat=True).distinct())
    annee_exercice_list.update(AttestationPriseEnCharge.objects.values_list('Annee_exercice', flat=True).distinct())
    annee_exercice_list.update(OrdreSortie.objects.values_list('Annee_exercice', flat=True).distinct())

    # Valeurs sélectionnées par l'utilisateur
    selected_nomenclature = request.GET.get('nomenclature')
    selected_annee_exercice = request.GET.get('annee_exercice')

    return {
        'nomenclature_list': sorted(nomenclature_list),  # Optionnel: trier la liste
        'annee_exercice_list': sorted(annee_exercice_list),  # Optionnel: trier la liste
        'selected_nomenclature': selected_nomenclature,
        'selected_annee_exercice': selected_annee_exercice,
    }
