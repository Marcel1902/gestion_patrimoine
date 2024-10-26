from django.shortcuts import render

# Create your views here.

#importation de toutes les bibliothèques nécessaire
from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth import login, logout, authenticate
from django.contrib.auth.decorators import login_required
from django.http import HttpResponseRedirect, HttpResponse
from django.db.models import Sum, F, Value
from django.urls import reverse
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from django.http import HttpResponse
import io
from django.contrib import messages
from openpyxl import Workbook
from .models import PVRecensement, Inventaire, EtatApreciatif, OrdreEntree, OrdreSortie, AttestationPriseEnCharge, Region, Service
# C'est ici qu'on va crée tout les fonctions


def ajout_region(request):
    if request.method == "POST":
        Nom = request.POST.get("nom")
        Region.objects.create(Nom=Nom)
        return redirect('afficher_regions')
    return render(request, "patrimoine/ajout_region.html")

def ajout_services(request, region_id):
    region = get_object_or_404(Region, id=region_id)
    if request.method == "POST":
        Nom = request.POST.get("nom")
        Service.objects.create(Region=region, Nom=Nom)  # Utilisation de l'objet Region
        return redirect(reverse('afficher_services', args=[region_id]))
    return render(request, 'patrimoine/ajout_services.html', {'region': region})

def afficher_regions(request):
    regions = Region.objects.all()
    return render(request, 'patrimoine/regions.html', {'regions': regions})

def afficher_services(request, region_id):
    region = get_object_or_404(Region, id=region_id)
    services = Service.objects.filter(Region=region)  # Utilisation du champ 'Region'
    return render(request, 'patrimoine/services.html', {'region': region, 'services': services})


#Page d'accueil, veux-dire, dès qu'on demarre l'application, c'est cette vue qu'i s'affichera en premier, notament c'est un Login
def home(request):
    if request.method == "POST":
        username = request.POST.get("nom")
        password = request.POST.get("password")
        user = authenticate(username=username, password=password)

        if user:
            login(request, user)
            return redirect('afficher_regions')
        else:
            messages.error(request, "Nom d'utilisateur ou mot de passe incorrect")
    return render(request, "patrimoine/home.html")

#fonction pour la deconnexion, on utilise @login_required pour interdire à tout personne d'y acceder sans se connecter
@login_required
def Deconnexion(request):
    logout(request)
    return redirect('home')


#fonction d'ajout d'un procès verbal de recensement
@login_required
def ajout_pv_recensement(request, service_id):
    service = get_object_or_404(Service, id=service_id)
    if request.method == "POST":
        Annee_exercice = request.POST.get("exercice")
        Nomenclature = request.POST.get("Nomenclature")
        Designation_materiels = request.POST.get("Designation_materiels")
        Especes_unites = request.POST.get("especes")
        Prix_unites = request.POST.get("prix")
        Quantites_d_apres_ecriture = request.POST.get("quantite_ecriture")
        Quantites_par_recensement = request.POST.get("quantite_recensement")
        Quantites_execedent_par_article = request.POST.get("excedent_article")
        Quantites_deficient_par_article = request.POST.get("deficient_article")
        valeurs_excedents_par_article = request.POST.get("valeurs_excedents_par_article")
        valeurs_excedents_par_nomenclature = request.POST.get("valeurs_excedents_par_nomenclature")
        valeurs_deficits_par_article = request.POST.get("valeurs_deficits_par_article")
        valeurs_des_deficits_par_nomenclature = request.POST.get("valeurs_des_deficits_par_nomenclature")
        valeurs_des_existants = request.POST.get("valeurs_des_existants")
        Observations = request.POST.get("Observations")
        
        PVRecensement.objects.create(
            Service=service,
            Annee_exercice=Annee_exercice,
            Nomenclature=Nomenclature,
            Designation_materiels=Designation_materiels,
            Especes_unites=Especes_unites,
            Prix_unites=Prix_unites,
            Quantites_d_apres_ecriture=Quantites_d_apres_ecriture,
            Quantites_par_recensement=Quantites_par_recensement,
            Quantites_execedent_par_article=Quantites_execedent_par_article,
            Quantites_deficient_par_article=Quantites_deficient_par_article,
            valeurs_excedents_par_article=valeurs_excedents_par_article,
            valeurs_excedents_par_nomenclature=valeurs_excedents_par_nomenclature,
            valeurs_deficits_par_article=valeurs_deficits_par_article,
            valeurs_des_deficits_par_nomenclature=valeurs_des_deficits_par_nomenclature,
            valeurs_des_existants=valeurs_des_existants,
            Observations=Observations
        )
        
        return redirect(reverse('pv_recensement', args=[service_id]))
    return render(request, "patrimoine/ajout_pv_recensement.html", {"service": service})



#fonction d'ajout d'un etat appreciatif
@login_required
def ajout_etat_apreciatif(request, service_id):
    service = get_object_or_404(Service, id=service_id)
    if request.method == "POST":
        Annee_exercice = request.POST.get("Annee_exercice")
        Nomenclature = request.POST.get("Nomenclature")
        Numero_du_piece_justificative = request.POST.get("Numero_du_piece_justificative")
        Date_du_piece_justificative = request.POST.get("Date_du_piece_justificative")
        Designations_sommaire_des_operations = request.POST.get("Designations_sommaire_des_operations")
        Charge = request.POST.get("Charge")
        Decharge = request.POST.get("Decharge")
        
        EtatApreciatif.objects.create(Service=service,
                                      Annee_exercice=Annee_exercice,
                                      Nomenclature=Nomenclature,
                                      Numero_du_piece_justificative=Numero_du_piece_justificative,
                                      Date_du_piece_justificative=Date_du_piece_justificative,
                                      Designations_sommaire_des_operations=Designations_sommaire_des_operations,
                                      Charge=Charge,
                                      Decharge=Decharge)
        return redirect(reverse('etat_appreciatif', args=[service_id]))
    return render(request, "patrimoine/ajout_etat_apreciatif.html", {"service": service})


#fonction d'ajout d'un inventaire
@login_required
def ajout_inventaire(request, service_id):
    service = get_object_or_404(Service, id=service_id)
    if request.method == "POST":
        Annee_exercice = request.POST.get("exercice")
        Nomenclature = request.POST.get("Nomenclature")
        Numero_folio_grand_livre = request.POST.get("Numero_folio_grand_livre")
        Designation_materiels = request.POST.get("Designation_materiels")
        Especes_des_unites = request.POST.get("Especes_des_unites")
        Prix_de_l_unite = request.POST.get("Prix_de_l_unite")
        Quantite_existant_1er_janvier = request.POST.get("Quantite_existant_1er_janvier")
        Quantite_entree_pendant_l_annee = request.POST.get("Quantite_entree_pendant_l_annee")
        Quantite_sortie_pendant_l_annee = request.POST.get("Quantite_sortie_pendant_l_annee")
        Quantite_reste_31_decembre = request.POST.get("Quantite_reste_31_decembre")
        Decompte = request.POST.get("Decompte")
        Observation = request.POST.get("Observation")
        
        Inventaire.objects.create(Service=service, 
                                  Annee_exercice=Annee_exercice,
                                  Nomenclature=Nomenclature,
                                  Numero_folio_grand_livre=Numero_folio_grand_livre,
                                  Designation_materiels=Designation_materiels,
                                  Especes_des_unites=Especes_des_unites,
                                  Prix_de_l_unite=Prix_de_l_unite,
                                  Quantite_existant_1er_janvier=Quantite_existant_1er_janvier,
                                  Quantite_entree_pendant_l_annee=Quantite_entree_pendant_l_annee,
                                  Quantite_sortie_pendant_l_annee=Quantite_sortie_pendant_l_annee,
                                  Quantite_reste_31_decembre=Quantite_reste_31_decembre,
                                  Decompte=Decompte,
                                  Observation=Observation)
        return redirect(reverse('inventaire', args=[service_id]))
    return render(request, "patrimoine/ajout_inventaire.html", {"service":service})


#fonction d'ajout d'un ordre d'entrée
@login_required
def ajout_ordre_entree(request, service_id):
    service = get_object_or_404(Service, id=service_id)
    if request.method == "POST":
        Annee_exercice = request.POST.get("Annee_exercice")
        Numero_folio_du_grandlivre = request.POST.get("Numero_folio_du_grandlivre")
        Nomenclature = request.POST.get("Nomenclature")
        Designation_des_matieres_et_objets = request.POST.get("Designation_des_matieres_et_objets")
        Especes_des_unites = request.POST.get("Especes_des_unites")
        Quantites = request.POST.get("Quantites")
        Prix_unite = request.POST.get("Prix_unite")
        Valeurs_partielles = request.POST.get("Valeurs_partielles")
        Valeurs_par_numero_nomenclature = request.POST.get("Valeurs_par_numero_nomenclature")
        Numero_piece_justificative_sortie_correspondante = request.POST.get("Numero_piece_justificative_sortie_correspondante")
        
        OrdreEntree.objects.create(Service=service,
                                   Annee_exercice=Annee_exercice,
                                   Numero_folio_du_grandlivre=Numero_folio_du_grandlivre,
                                   Nomenclature=Nomenclature,
                                   Designation_des_matieres_et_objets=Designation_des_matieres_et_objets,
                                   Especes_des_unites=Especes_des_unites,
                                   Quantites=Quantites,
                                   Prix_unite=Prix_unite,
                                   Valeurs_partielles=Valeurs_partielles,
                                   Valeurs_par_numero_nomenclature=Valeurs_par_numero_nomenclature,
                                   Numero_piece_justificative_sortie_correspondante=Numero_piece_justificative_sortie_correspondante)
        
        return redirect(reverse('ordre_entree', args=[service_id]))
    return render(request, "patrimoine/ajout_ordre_entree.html", {"service": service})


#fonction d'ajout d'un attestation de prise en charge
@login_required
def ajout_attestation_prise_en_charge(request, service_id):
    service = get_object_or_404(Service, id=service_id)
    if request.method == "POST":
        Annee_exercice = request.POST.get("Annee_exercice")
        Nomenclature = request.POST.get("Nomenclature")
        Designation_des_matieres_et_objets = request.POST.get("Designation_des_matieres_et_objets")
        Especes_des_unites = request.POST.get("Especes_des_unites")
        Quantite = request.POST.get("Quantite")
        Prix_unite = request.POST.get("Prix_unite")
        Montant = request.POST.get("Montant")
        Observations = request.POST.get("Observations")
        
        AttestationPriseEnCharge.objects.create(Designation_des_matieres_et_objets=Designation_des_matieres_et_objets,
                                                Nomenclature=Nomenclature,
                                                Annee_exercice=Annee_exercice,
                                                Service=service,
                                                Especes_des_unites=Especes_des_unites,
                                                Quantite=Quantite,
                                                Prix_unite=Prix_unite,
                                                Montant=Montant,
                                                Observations=Observations)
        
        return redirect(reverse('attestation', args=[service_id]))
    return render(request, "patrimoine/ajout_attestation.html", {"service":service})


#fonction d'ajout d'un ordre de sortie
@login_required
def ajout_ordre_sortie(request, service_id):
    service = get_object_or_404(Service, id=service_id)
    if request.method == "POST":
        Annee_exercice = request.POST.get("Annee_exercice")
        Numero_folio_du_grandlivre = request.POST.get("Numero_folio_du_grandlivre")
        Nomenclature = request.POST.get("Nomenclature")
        Designation_des_matieres_et_objets = request.POST.get("Designation_des_matieres_et_objets")
        Especes_des_unites = request.POST.get("Especes_des_unites")
        Quantites = request.POST.get("Quantites")
        Prix_unite = request.POST.get("Prix_unite")
        Valeurs_partielles = request.POST.get("Valeurs_partielles")
        Valeurs_par_numero_nomenclature = request.POST.get("Valeurs_par_numero_nomenclature")
        Numero_piece_justificative_sortie_correspondante = request.POST.get("Numero_piece_justificative_sortie_correspondante")
        
        OrdreSortie.objects.create(Annee_exercice=Annee_exercice,
                                   Service=service,
                                   Numero_folio_du_grandlivre=Numero_folio_du_grandlivre,
                                   Nomenclature=Nomenclature,
                                   Designation_des_matieres_et_objets=Designation_des_matieres_et_objets,
                                   Especes_des_unites=Especes_des_unites,
                                   Quantites=Quantites,
                                   Prix_unite=Prix_unite,
                                   Valeurs_partielles=Valeurs_partielles,
                                   Valeurs_par_numero_nomenclature=Valeurs_par_numero_nomenclature,
                                   Numero_piece_justificative_sortie_correspondante=Numero_piece_justificative_sortie_correspondante)
        
        return redirect(reverse('ordre_sortie', args=[service_id]))
    return render(request, "patrimoine/ajout_ordre_sortie.html", {"service":service})


#fonction d'affichage du procès verbal de recensement
@login_required
def PVRecensement_table(request, service_id):
    # Récupérer tous les objets PVRecensement
    service = get_object_or_404(Service, id=service_id)
    pvrecensements = PVRecensement.objects.filter(Service=service)

    # Récupérer les valeurs des filtres depuis les paramètres GET
    selected_nomenclature = request.GET.get('nomenclature')
    selected_annee_exercice = request.GET.get('annee_exercice')

    # Filtrer les données en fonction des filtres sélectionnés
    if selected_nomenclature:
        pvrecensements = pvrecensements.filter(Nomenclature=selected_nomenclature)
    if selected_annee_exercice:
        pvrecensements = pvrecensements.filter(Annee_exercice=selected_annee_exercice)

    # Calculer les récapitulatifs en fonction des filtres appliqués
    recap = pvrecensements.values('Nomenclature').annotate(
        total_existants=Sum('valeurs_des_existants'),
        total_excedents=Sum('valeurs_excedents_par_article'),
        total_deficits=Sum('valeurs_deficits_par_article')
    ).order_by('Nomenclature')

    # Calculer le total général en fonction des données filtrées
    Total_general = {
        'total_valeurs_des_existants': recap.aggregate(Sum('total_existants'))['total_existants__sum'] or 0,
        'total_valeurs_excedents_par_article': recap.aggregate(Sum('total_excedents'))['total_excedents__sum'] or 0,
        'total_valeurs_deficits_par_article': recap.aggregate(Sum('total_deficits'))['total_deficits__sum'] or 0,
    }

    # Renvoyer les données au template
    context = {
        "pvrecensements": pvrecensements,
        "Total_general": Total_general,
        "recap": recap,
        "selected_nomenclature": selected_nomenclature,
        "selected_annee_exercice": selected_annee_exercice,
        "service": service,
    }

    return render(request, "patrimoine/pvrecensement.html", context)


#fonction d'affichage de l'inventaire
@login_required
def Inventaire_table(request, service_id):
    # Récupérer tous les objets Inventaire
    service = get_object_or_404(Service, id=service_id)
    inventaires = Inventaire.objects.filter(Service=service)

    # Récupérer les valeurs des filtres depuis les paramètres GET
    selected_nomenclature = request.GET.get('nomenclature')
    selected_annee_exercice = request.GET.get('annee_exercice')

    # Filtrer les données en fonction des filtres sélectionnés
    if selected_nomenclature:
        inventaires = inventaires.filter(Nomenclature=selected_nomenclature)
    if selected_annee_exercice:
        inventaires = inventaires.filter(Annee_exercice=selected_annee_exercice)

    # Calculer les récapitulatifs en fonction des données filtrées
    recap = inventaires.values('Nomenclature').annotate(
        decompte=Sum('Decompte'),
    ).order_by('Nomenclature')

    # Calculer le total général en fonction des données filtrées
    Total_general = {
        'total_decompte': recap.aggregate(Sum('decompte'))['decompte__sum'] or 0,
    }

    # Renvoyer les données au template
    context = {
        "inventaires": inventaires,
        "recap": recap,
        "Total_general": Total_general,
        "selected_nomenclature": selected_nomenclature,
        "selected_annee_exercice": selected_annee_exercice,
        "service": service,
    }

    return render(request, "patrimoine/inventaire.html", context)



@login_required
def recapitulatif_inventaire(request, service_id):
    service = get_object_or_404(Service, id=service_id)
    inventaire = Inventaire.objects.filter(Service=service)

    # Récupérer les valeurs des filtres depuis les paramètres GET
    selected_nomenclature = request.GET.get('nomenclature')
    selected_annee_exercice = request.GET.get('annee_exercice')

    # Calculer les totaux par nomenclature
    recapitulatif = (
        inventaire
        .values('Nomenclature', 'Annee_exercice')
        .annotate(
            total_valeur=Sum('Decompte'),
            total_prix_janvier=Sum(F('Quantite_existant_1er_janvier') * F('Prix_de_l_unite')),
            total_entrees_annee=Sum(F('Quantite_entree_pendant_l_annee') * F('Prix_de_l_unite')),  # Utilisation de la relation
            total_existant_et_entrees=Sum(
                (F('Quantite_existant_1er_janvier') + F('Quantite_entree_pendant_l_annee')) * F('Prix_de_l_unite')),
            total_sorties=Sum(F('Quantite_sortie_pendant_l_annee') * F('Prix_de_l_unite')),
            total_reste=Sum(F('Quantite_reste_31_decembre') * F('Prix_de_l_unite'))
        )
        .order_by('Nomenclature')
    )

    # Appliquer les filtres en fonction des filtres sélectionnés
    if selected_nomenclature:
        recapitulatif = recapitulatif.filter(Nomenclature=selected_nomenclature)
    if selected_annee_exercice:
        recapitulatif = recapitulatif.filter(Annee_exercice=selected_annee_exercice)

    # Calculer le total général en itérant sur les résultats du récapitulatif
    total_general = {
        'total_nomenclature': sum(item['total_valeur'] for item in recapitulatif),
        'total_prix_janvier': sum(item['total_prix_janvier'] for item in recapitulatif),
        'total_entrees_annee': sum(item['total_entrees_annee'] or 0 for item in recapitulatif),
        'total_existant_et_entrees': sum(item['total_existant_et_entrees'] for item in recapitulatif),
        'total_sorties': sum(item['total_sorties'] or 0 for item in recapitulatif),
        'total_reste': sum(item['total_reste'] for item in recapitulatif)
    }

    # Préparer le contexte
    context = {
        'recapitulatif': recapitulatif,
        'total_general': total_general,
        'service': service,
        'selected_nomenclature': selected_nomenclature,
        'selected_annee_exercice': selected_annee_exercice,
    }
    return render(request, 'patrimoine/recapitulatif.html', context)



#fonction d'affichage de l'etat appreciatif
@login_required
def EtatApreciatif_table(request, service_id):
    # Récupérer tous les objets EtatApreciatif
    service = get_object_or_404(Service, id=service_id)
    etatappreciatifs = EtatApreciatif.objects.filter(Service=service)

    # Récupérer les valeurs des filtres depuis les paramètres GET
    selected_nomenclature = request.GET.get('nomenclature')
    selected_annee_exercice = request.GET.get('annee_exercice')

    # Filtrer les données en fonction des filtres sélectionnés
    if selected_nomenclature:
        etatappreciatifs = etatappreciatifs.filter(Nomenclature=selected_nomenclature)
    if selected_annee_exercice:
        etatappreciatifs = etatappreciatifs.filter(Annee_exercice=selected_annee_exercice)

    # Calculer les récapitulatifs en fonction des données filtrées
    recap = etatappreciatifs.values('Nomenclature').annotate(
        charge=Sum('Charge'),
        decharge=Sum('Decharge'),
    ).order_by('Nomenclature')

    # Calculer le total général en fonction des données filtrées
    Total_general = {
        'total_charge': recap.aggregate(Sum('charge'))['charge__sum'] or 0,
        'total_decharge': recap.aggregate(Sum('decharge'))['decharge__sum'] or 0,
    }

    # Renvoyer les données au template
    context = {
        "etatappreciatifs": etatappreciatifs,
        "recap": recap,
        "Total_general": Total_general,
        "selected_nomenclature": selected_nomenclature,
        "selected_annee_exercice": selected_annee_exercice,
        "service": service,
    }

    return render(request, "patrimoine/etat_appreciatif.html", context)

#fonction d'affichage de l'ordre entree
@login_required
def OrdreEntree_table(request, service_id):
    # Récupérer tous les objets OrdreEntree
    service = get_object_or_404(Service, id=service_id)
    ordreEntrees = OrdreEntree.objects.filter(Service=service)

    # Récupérer les valeurs des filtres depuis les paramètres GET
    selected_nomenclature = request.GET.get('nomenclature')
    selected_annee_exercice = request.GET.get('annee_exercice')

    # Filtrer les données en fonction des filtres sélectionnés
    if selected_nomenclature:
        ordreEntrees = ordreEntrees.filter(Nomenclature=selected_nomenclature)
    if selected_annee_exercice:
        ordreEntrees = ordreEntrees.filter(Annee_exercice=selected_annee_exercice)

    # Calculer les récapitulatifs en fonction des données filtrées
    recap = ordreEntrees.values('Nomenclature').annotate(
        Valeurs_par_numero_nomenclature=Sum('Valeurs_par_numero_nomenclature')
    ).order_by('Nomenclature')

    # Calculer le total général en fonction des données filtrées
    Total_general = {
        'total_valeurs_par_nomenclature': recap.aggregate(Sum('Valeurs_par_numero_nomenclature'))['Valeurs_par_numero_nomenclature__sum'] or 0,
    }

    # Renvoyer les données au template
    context = {
        "ordreEntrees": ordreEntrees,
        "recap": recap,
        "Total_general": Total_general,
        "selected_nomenclature": selected_nomenclature,
        "selected_annee_exercice": selected_annee_exercice,
        "service": service,
    }

    return render(request, "patrimoine/ordre_entree.html", context)

#fonction d'affichage de l'ordre de sortie
@login_required
def OrdreSortie_table(request, service_id):
    # Récupérer tous les objets OrdreSortie
    service = get_object_or_404(Service, id=service_id)
    ordresorties = OrdreSortie.objects.filter(Service=service)

    # Récupérer les valeurs des filtres depuis les paramètres GET
    selected_nomenclature = request.GET.get('nomenclature')
    selected_annee_exercice = request.GET.get('annee_exercice')

    # Filtrer les données en fonction des filtres sélectionnés
    if selected_nomenclature:
        ordresorties = ordresorties.filter(Nomenclature=selected_nomenclature)
    if selected_annee_exercice:
        ordresorties = ordresorties.filter(Annee_exercice=selected_annee_exercice)

    # Calculer les récapitulatifs en fonction des données filtrées
    recap = ordresorties.values('Nomenclature').annotate(
        Valeurs_par_numero_nomenclature=Sum('Valeurs_par_numero_nomenclature')
    ).order_by('Nomenclature')

    # Calculer le total général en fonction des données filtrées
    Total_general = {
        'total_valeurs_par_nomenclature': recap.aggregate(Sum('Valeurs_par_numero_nomenclature'))['Valeurs_par_numero_nomenclature__sum'] or 0,
    }

    # Renvoyer les données au template
    context = {
        "ordresorties": ordresorties,
        "Total_general": Total_general,
        "recap": recap,
        "selected_nomenclature": selected_nomenclature,
        "selected_annee_exercice": selected_annee_exercice,
        "service": service,
    }

    return render(request, "patrimoine/ordre_sortie.html", context)

#fonction d'affichage de l'attestation de prise en charge
@login_required
def AttestationPriseEnCharge_table(request, service_id):
    service = get_object_or_404(Service, id=service_id)
    attestationpriseEncharges = AttestationPriseEnCharge.objects.filter(Service=service)
    
    # Récupérer les valeurs des filtres depuis les paramètres GET
    selected_nomenclature = request.GET.get('nomenclature')
    selected_annee_exercice = request.GET.get('annee_exercice')

    # Filtrer les données en fonction des filtres sélectionnés
    if selected_nomenclature:
        attestationpriseEncharges = attestationpriseEncharges.filter(Nomenclature=selected_nomenclature)
    if selected_annee_exercice:
        attestationpriseEncharges = attestationpriseEncharges.filter(Annee_exercice=selected_annee_exercice)
    recap = AttestationPriseEnCharge.objects.values('Designation_des_matieres_et_objets').annotate(total_general = Sum('Montant'),
                                                                  ).order_by('Designation_des_matieres_et_objets')
    Total_general = {
        'total_general': recap.aggregate(Sum('total_general'))['total_general__sum'] or 0,
    }
    return render(request, "patrimoine/attestation.html", {"attestationpriseEncharges": attestationpriseEncharges, "Total_general": Total_general, "recap": recap, "service": service})


#mettre le mot nomenclature dans un dictionnaire pour eviter le problème dans le filtre
FIELD_MAPPING = {
    'nomenclature': 'Nomenclature',
    'annee_exercice': 'Annee_exercice',
    # Ajoutez d'autres correspondances si nécessaire
}

def export_data(request, model_class, fields, filename_prefix, service_id):
    filters = {k: request.GET.get(k) for k in request.GET.keys() if k not in ['format']}
    format_export = request.GET.get('format', '').lower()  # Normaliser en minuscules

    # Debugging: Afficher les filtres et le format
    print(f"Filtres : {filters}")
    print(f"Format demandé : {format_export}")

    # Récupérer le service spécifique par son ID
    service = get_object_or_404(Service, id=service_id)

    # Récupérer toutes les données pour le service spécifié
    data_queryset = model_class.objects.filter(Service=service)

    # Appliquer d'autres filtres
    for key, value in filters.items():
        if value and value.lower() != 'none':  # Ignorez 'None'
            correct_key = FIELD_MAPPING.get(key, key)
            if hasattr(model_class, correct_key):
                print(f"Filtrage avec : {{{correct_key}: {value}}}")  # Log de filtrage
                data_queryset = data_queryset.filter(**{correct_key: value})
            else:
                return HttpResponse(f"Le champ '{key}' n'existe pas dans le modèle.", status=400)

    # Exporter selon le format choisi
    if format_export == 'pdf':
        return export_to_pdf(data_queryset, fields, filename_prefix)
    elif format_export == 'excel':
        return export_to_excel(data_queryset, fields, filename_prefix)
    else:
        return HttpResponse(f"Format non supporté: '{format_export}'", status=400)


def export_to_pdf(data_queryset, fields, filename_prefix):
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = f'attachment; filename="{filename_prefix}.pdf"'
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)

    # Préparer les éléments du PDF
    elements = []

    # Ajouter en-tête
    title_style = ParagraphStyle(name='TitleStyle', fontName='Helvetica-Bold', fontSize=14, alignment=1)
    elements.append(Paragraph("REPOBLIKAN'I MADAGASIKARA", title_style))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("Fitiavana-Tanindrazana-Fandrosoana", title_style))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("************", title_style))
    elements.append(Spacer(1, 20))

    # Ajouter les informations de budget (à remplacer par les valeurs réelles)
    budget_style = ParagraphStyle(name='BudgetStyle', fontName='Helvetica', fontSize=10, alignment=1)
    elements.append(Paragraph("BUDGET: ", budget_style))
    elements.append(Paragraph("EXERCICE: ", budget_style))
    elements.append(Paragraph("IMPUTATION ADMINISTRATIVE: ", budget_style))
    elements.append(Paragraph("CHAPITRE:  - ARTICLE: - PARAGRAPHE: ", budget_style))
    elements.append(Spacer(1, 20))

    # Préparer les données pour le tableau
    data = [fields]  # En-tête du tableau

    # Ajouter les données des objets filtrés
    for item in data_queryset:
        row = [getattr(item, field) for field in fields]
        data.append(row)

    # Créer le tableau PDF
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),  # Correction pour dessiner les bordures du tableau
    ]))

    # Ajouter le tableau au document PDF
    elements.append(table)
    doc.build(elements)

    # Récupérer le contenu du PDF
    pdf = buffer.getvalue()
    buffer.close()
    response.write(pdf)
    return response


def export_to_excel(data_queryset, fields, filename_prefix):
    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = f'attachment; filename="{filename_prefix}.xlsx"'
    wb = Workbook()
    ws = wb.active
    ws.title = "Inventaire"

    # En-têtes personnalisées
    ws.append(['REPOBLIKAN\'I MADAGASIKARA', '', '', '', '', '', '', '', '', '', '', '', ''])
    ws.append(['Fitiavana-Tanindrazana-Fandrosoana', '', '', '', '', '', '', '', '', '', '', '', ''])
    ws.append(['************', '', '', '', '', '', '', '', '', '', '', '', ''])

    # Ajouter les informations de budget (à remplacer par les valeurs réelles)
    ws.append(['BUDGET', '[Budget]'])
    ws.append(['EXERCICE', '[Exercice]'])
    ws.append(['IMPUTATION ADMINISTRATIVE', '[Imputation]'])
    ws.append(['CHAPITRE', '[Chapitre]', 'ARTICLE', '[Article]', 'PARAGRAPHE', '[Paragraphe]'])
    ws.append([''])  # Ajouter une ligne vide

    # En-têtes du tableau
    ws.append(fields)

    # Ajouter les données
    for item in data_queryset:
        row = [getattr(item, field) for field in fields]
        ws.append(row)

    # Enregistrer le fichier Excel dans la réponse
    wb.save(response)
    return response


def export_pvrecensement(request, service_id):
    fields = [
        'Annee_exercice', 'Nomenclature', 'Designation_materiels', 'Especes_unites',
        'Prix_unites', 'Quantites_d_apres_ecriture', 'Quantites_par_recensement',
        'Quantites_execedent_par_article', 'Quantites_deficient_par_article',
        'valeurs_excedents_par_article', 'valeurs_excedents_par_nomenclature',
        'valeurs_deficits_par_article', 'valeurs_des_deficits_par_nomenclature',
        'valeurs_des_existants', 'Observations'
    ]
    filters = {k: request.GET.get(k) for k in request.GET.keys() if k not in ['format']}
    # Récupérer le service spécifique par son ID
    service = get_object_or_404(Service, id=service_id)
    # Filtrer les données en fonction du service_id
    data_queryset = PVRecensement.objects.filter(Service=service)

    # Vérifier le format demandé
    format_export = request.GET.get('format')
    if format_export == 'pdf':
        return export_to_pdf(data_queryset, fields, 'pvrecensement')
    elif format_export == 'excel':
        return export_to_excel(data_queryset, fields, 'pvrecensement')
    else:
        return HttpResponse(f"Format non supporté : '{format_export}'", status=400)


#fonction pour l'exportation du données de la table Etat Apreciatif
def export_etat_apreciatif(request, service_id):
    #ajouter toutes les champs de la table dans fields
    fields = [
        'Annee_exercice', 'Nomenclature', 'Numero_du_piece_justificative',
        'Date_du_piece_justificative', 'Designations_sommaire_des_operations',
        'Charge', 'Decharge'
    ]

    # Récupérer le service spécifique par son ID
    service = get_object_or_404(Service, id=service_id)
    # Filtrer les données en fonction du service_id
    data_queryset = EtatApreciatif.objects.filter(Service=service)
    # Vérifier le format demandé
    format_export = request.GET.get('format')
    if format_export == 'pdf':
        return export_to_pdf(data_queryset, fields, 'Etat appreciatif')
    elif format_export == 'excel':
        return export_to_excel(data_queryset, fields, 'Etat appreciatif')
    else:
        return HttpResponse(f"Format non supporté : '{format_export}'", status=400)


#fonction pour l'exportation du données de la table Inventaire
def export_inventaire(request, service_id):
    #ajouter toutes les champs de la table dans fields
    fields = [
        'Annee_exercice', 'Nomenclature', 'Numero_folio_grand_livre',
        'Designation_materiels', 'Especes_des_unites', 'Prix_de_l_unite',
        'Quantite_existant_1er_janvier', 'Quantite_entree_pendant_l_annee',
        'Quantite_sortie_pendant_l_annee', 'Quantite_reste_31_decembre',
        'Decompte', 'Observation'
    ]
    # Récupérer le service spécifique par son ID
    service = get_object_or_404(Service, id=service_id)
    # Filtrer les données en fonction du service_id
    data_queryset = Inventaire.objects.filter(Service=service)
    # Vérifier le format demandé
    format_export = request.GET.get('format')
    if format_export == 'pdf':
        return export_to_pdf(data_queryset, fields, 'Inventaire')
    elif format_export == 'excel':
        return export_to_excel(data_queryset, fields, 'Inventaire')
    else:
        return HttpResponse(f"Format non supporté : '{format_export}'", status=400)

#fonction pour l'exportation du données de la table Ordre entrée
def export_ordre_entree(request, service_id):
    #ajouter toutes les champs de la table dans fields
    fields = [
        'Annee_exercice', 'Numero_folio_du_grandlivre', 'Nomenclature',
        'Designation_des_matieres_et_objets', 'Especes_des_unites', 'Quantites',
        'Prix_unite', 'Valeurs_partielles', 'Valeurs_par_numero_nomenclature',
        'Numero_piece_justificative_sortie_correspondante'
    ]
    # Récupérer le service spécifique par son ID
    service = get_object_or_404(Service, id=service_id)
    # Filtrer les données en fonction du service_id
    data_queryset = OrdreEntree.objects.filter(Service=service)
    # Vérifier le format demandé
    format_export = request.GET.get('format')
    if format_export == 'pdf':
        return export_to_pdf(data_queryset, fields, 'ordre_entree')
    elif format_export == 'excel':
        return export_to_excel(data_queryset, fields, 'ordre_entree')
    else:
        return HttpResponse(f"Format non supporté : '{format_export}'", status=400)

#fonction pour l'exportation du données de la table Attestation prise en charge
def export_attestation_prise_en_charge(request, service_id):
    #ajouter toutes les champs de la table dans fields
    fields = [
        'Designation_des_matieres_et_objets', 'Especes_des_unites',
        'Quantite', 'Prix_unite', 'Montant', 'Observations'
    ]
    # Récupérer le service spécifique par son ID
    service = get_object_or_404(Service, id=service_id)
    # Filtrer les données en fonction du service_id
    data_queryset = AttestationPriseEnCharge.objects.filter(Service=service)
    # Vérifier le format demandé
    format_export = request.GET.get('format')
    if format_export == 'pdf':
        return export_to_pdf(data_queryset, fields, 'attestation_prise_en_charge')
    elif format_export == 'excel':
        return export_to_excel(data_queryset, fields, 'attestation_prise_en_charge')
    else:
        return HttpResponse(f"Format non supporté : '{format_export}'", status=400)

#fonction pour l'exportation du données de la table Ordre de sortie
def export_ordre_sortie(request, service_id):
    #ajouter toutes les champs de la table dans fields
    fields = [
        'Annee_exercice', 'Numero_folio_du_grandlivre', 'Nomenclature',
        'Designation_des_matieres_et_objets', 'Especes_des_unites', 'Quantites',
        'Prix_unite', 'Valeurs_partielles', 'Valeurs_par_numero_nomenclature',
        'Numero_piece_justificative_sortie_correspondante'
    ]
    # Récupérer le service spécifique par son ID
    service = get_object_or_404(Service, id=service_id)
    # Filtrer les données en fonction du service_id
    data_queryset = OrdreSortie.objects.filter(Service=service)
    # Vérifier le format demandé
    format_export = request.GET.get('format')
    if format_export == 'pdf':
        return export_to_pdf(data_queryset, fields, 'ordre_sortie')
    elif format_export == 'excel':
        return export_to_excel(data_queryset, fields, 'ordre_sortie')
    else:
        return HttpResponse(f"Format non supporté : '{format_export}'", status=400)
