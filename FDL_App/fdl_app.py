# -*- coding: utf-8 -*-
"""
FDL - Programme des dimanches
Application Windows (Python + Tkinter) : groupes, membres, répartition équilibrée des rôles,
export Word au même format que le programme actuel.

Lancer :  py fdl_app.py      (ou double-clic sur Lancer_FDL.bat)
"""
import json
import os
import re
import sys
import webbrowser
import datetime as dt
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fdl_core as core            # noqa: E402
import fdl_telechargements as dl   # noqa: E402
import excel_vers_word as ew       # noqa: E402

NAVY = "#17067A"        # bleu du logo FJKM
NAVY_CLAIR = "#3A2F9E"
FOND = "#EEF1F9"        # fond de l'interface
FOND_ONGLET = "#D5DAEE"
TEXTE = "#1B1B3A"
BORDURE = "#B9C1DD"
POLICE = "Segoe UI"
GITHUB_UTILISATEUR = "patrickandriamalala-art"      # proposé par défaut dans les fenêtres GitHub


def ressource(*parts):
    """Chemin d'un fichier fourni avec l'application (fonctionne aussi dans le .exe)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, *parts)


JOURS = ["lun.", "mar.", "mer.", "jeu.", "ven.", "sam.", "dim."]
ROLES_COURTS = {1: "VVK VKT", 2: "VKT 1", 3: "VKT 2", 4: "VKT 3", 5: "Rôle 5"}


def chemin_config():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    dossier = os.path.join(base, "FDL")
    os.makedirs(dossier, exist_ok=True)
    return os.path.join(dossier, "config.json")


def lire_config():
    try:
        import json
        with open(chemin_config(), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def ecrire_config(cfg):
    try:
        import json
        with open(chemin_config(), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False)
    except Exception:
        pass


def fmt(d):
    if isinstance(d, str):
        d = core.s2d(d)
    return d.strftime("%d/%m/%Y")


class DialogSemestre(tk.Toplevel):
    """Choix du semestre à écrire dans le Word."""

    def __init__(self, parent, sem, annee):
        super().__init__(parent)
        self.title("Exporter en Word")
        self.configure(background=FOND)
        self.resultat = None
        self.resizable(False, False)
        self.transient(parent)
        f = ttk.Frame(self, padding=14)
        f.pack()
        ttk.Label(f, text="Le Word contient 6 mois (2 tableaux de 3 mois).").grid(row=0, column=0, columnspan=2, pady=(0, 8))
        ttk.Label(f, text="Semestre :").grid(row=1, column=0, sticky="e", padx=4, pady=3)
        self.v_sem = tk.StringVar(value=sem)
        ttk.Combobox(f, textvariable=self.v_sem, values=["S1", "S2"], width=6, state="readonly").grid(row=1, column=1, sticky="w")
        ttk.Label(f, text="S1 = janvier–juin    S2 = juillet–décembre", foreground="#555").grid(row=2, column=0, columnspan=2)
        ttk.Label(f, text="Année :").grid(row=3, column=0, sticky="e", padx=4, pady=3)
        self.v_an = tk.StringVar(value=str(annee))
        ttk.Spinbox(f, from_=2020, to=2100, textvariable=self.v_an, width=7).grid(row=3, column=1, sticky="w")
        b = ttk.Frame(f)
        b.grid(row=4, column=0, columnspan=2, pady=(12, 0))
        ttk.Button(b, text="Continuer…", command=self.ok).pack(side="left", padx=4)
        ttk.Button(b, text="Annuler", command=self.destroy).pack(side="left", padx=4)
        self.grab_set()
        self.wait_window()

    def ok(self):
        try:
            self.resultat = (self.v_sem.get(), int(self.v_an.get()))
        except ValueError:
            messagebox.showerror("Année", "Année invalide.", parent=self)
            return
        self.destroy()


class DialogAbsences(tk.Toplevel):
    """Périodes d'absence d'un membre."""

    def __init__(self, parent, titre, absences):
        super().__init__(parent)
        self.title(titre)
        self.configure(background=FOND)
        self.absences = [list(a) for a in absences]
        self.transient(parent)
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Pendant ces périodes, la personne n'est jamais choisie.").pack(anchor="w")
        self.lb = tk.Listbox(f, height=8, width=44, background="white", selectbackground=NAVY,
                             selectforeground="white", highlightthickness=1, highlightbackground=BORDURE,
                             relief="flat", font=(POLICE, 10))
        self.lb.pack(fill="both", expand=True, pady=6)
        g = ttk.Frame(f)
        g.pack(fill="x")
        ttk.Label(g, text="du").pack(side="left")
        self.e1 = ttk.Entry(g, width=11)
        self.e1.pack(side="left", padx=3)
        ttk.Label(g, text="au").pack(side="left")
        self.e2 = ttk.Entry(g, width=11)
        self.e2.pack(side="left", padx=3)
        ttk.Button(g, text="Ajouter", command=self.ajouter).pack(side="left", padx=3)
        ttk.Label(f, text="Format JJ/MM/AAAA", foreground="#555").pack(anchor="w")
        h = ttk.Frame(f)
        h.pack(fill="x", pady=(8, 0))
        ttk.Button(h, text="Supprimer la sélection", command=self.supprimer).pack(side="left")
        ttk.Button(h, text="Terminé", command=self.destroy).pack(side="right")
        self.rafraichir()
        self.grab_set()
        self.wait_window()

    def rafraichir(self):
        self.lb.delete(0, "end")
        for a, b in self.absences:
            self.lb.insert("end", f"du {fmt(a)} au {fmt(b)}")

    def ajouter(self):
        try:
            a, b = core.lire_date(self.e1.get()), core.lire_date(self.e2.get())
        except Exception:
            messagebox.showerror("Date", "Dates invalides (JJ/MM/AAAA).", parent=self)
            return
        if b < a:
            messagebox.showerror("Date", "La date de fin précède la date de début.", parent=self)
            return
        self.absences.append([core.d2s(a), core.d2s(b)])
        self.absences.sort()
        self.e1.delete(0, "end")
        self.e2.delete(0, "end")
        self.rafraichir()

    def supprimer(self):
        for i in sorted(self.lb.curselection(), reverse=True):
            del self.absences[i]
        self.rafraichir()


class DialogTelechargements(tk.Toplevel):
    """Crée la page de téléchargement + le QR code (à déposer sur un dépôt GitHub public)."""

    def __init__(self, parent, config):
        super().__init__(parent)
        self.title("Téléchargements pour les membres (QR code)")
        self.configure(background=FOND)
        self.transient(parent)
        self.geometry("820x640")
        self.fichiers = [dict(f) for f in config.get("fichiers", [])]
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        f.columnconfigure(1, weight=1)
        ttk.Label(f, text="Votre nom d'utilisateur GitHub :").grid(row=0, column=0, sticky="e", pady=3, padx=4)
        self.v_user = tk.StringVar(value=config.get("utilisateur") or GITHUB_UTILISATEUR)
        ttk.Entry(f, textvariable=self.v_user, width=30).grid(row=0, column=1, sticky="w")
        ttk.Label(f, text="Nom du dépôt PUBLIC :").grid(row=1, column=0, sticky="e", pady=3, padx=4)
        self.v_depot = tk.StringVar(value=config.get("depot", "fdl-telechargements"))
        ttk.Entry(f, textvariable=self.v_depot, width=30).grid(row=1, column=1, sticky="w")
        ttk.Label(f, text="Langue proposée au départ :").grid(row=2, column=0, sticky="e", pady=3, padx=4)
        self.v_langue = tk.StringVar(value="Malagasy" if config.get("langue_defaut") == "mg" else "Français")
        ttk.Combobox(f, textvariable=self.v_langue, values=["Français", "Malagasy"], state="readonly", width=12).grid(row=2, column=1, sticky="w")
        self.v_libre = tk.BooleanVar(value=bool(config.get("langue_libre", False)))
        ttk.Checkbutton(f, text="Laisser changer de langue (bouton Malagasy / Français)", variable=self.v_libre).grid(row=2, column=1, sticky="e")
        self.lbl_url = ttk.Label(f, text="", foreground=NAVY, font=(POLICE, 10, "bold"))
        self.lbl_url.grid(row=3, column=0, columnspan=2, sticky="w", padx=4, pady=(6, 2))
        for v in (self.v_user, self.v_depot):
            v.trace_add("write", lambda *a: self.maj_url())

        ttk.Label(f, text="Fichiers à proposer :").grid(row=4, column=0, columnspan=2, sticky="w", padx=4, pady=(8, 2))
        self.lb = tk.Listbox(f, height=9, background="white", selectbackground=NAVY, selectforeground="white",
                             highlightthickness=1, highlightbackground=BORDURE, relief="flat", font=(POLICE, 10))
        self.lb.grid(row=5, column=0, columnspan=2, sticky="nsew")
        f.rowconfigure(5, weight=1)
        b = ttk.Frame(f)
        b.grid(row=6, column=0, columnspan=2, sticky="w", pady=6)
        ttk.Button(b, text="Ajouter l'APK…", command=lambda: self.ajouter("apk")).pack(side="left", padx=2)
        ttk.Button(b, text="Ajouter la page du programme…", command=lambda: self.ajouter("page")).pack(side="left", padx=2)
        ttk.Button(b, text="Ajouter un autre fichier…", command=lambda: self.ajouter("autre")).pack(side="left", padx=2)
        ttk.Button(b, text="Retirer", command=self.retirer).pack(side="left", padx=8)

        ttk.Label(f, text="Dossier à créer :").grid(row=7, column=0, sticky="e", pady=3, padx=4)
        g = ttk.Frame(f)
        g.grid(row=7, column=1, sticky="we")
        self.v_dossier = tk.StringVar(value=config.get("dossier", os.path.join(os.path.expanduser("~"), "FDL_telechargements")))
        ttk.Entry(g, textvariable=self.v_dossier).pack(side="left", fill="x", expand=True)
        ttk.Button(g, text="Parcourir…", command=self.parcourir).pack(side="left", padx=4)
        ttk.Label(f, text="Le dossier contiendra la page, les fichiers, le QR code (image) et une feuille Word à imprimer.\n"
                          "Déposez ensuite TOUT son contenu sur le dépôt GitHub public (voir GUIDE_TELECHARGEMENTS_QR.txt).",
                  foreground="#555", justify="left").grid(row=8, column=0, columnspan=2, sticky="w", padx=4, pady=4)
        h = ttk.Frame(f)
        h.grid(row=9, column=0, columnspan=2, sticky="e", pady=(8, 0))
        ttk.Button(h, text="Créer la page et le QR code", command=self.creer).pack(side="left", padx=4)
        ttk.Button(h, text="Fermer", command=self.destroy).pack(side="left", padx=4)
        self.maj_url()
        self.rafraichir()

    def config(self):
        return {"utilisateur": self.v_user.get().strip(), "depot": self.v_depot.get().strip(),
                "langue_defaut": "mg" if self.v_langue.get() == "Malagasy" else "fr",
                "langue_libre": bool(self.v_libre.get()),
                "dossier": self.v_dossier.get().strip(), "fichiers": self.fichiers}

    def maj_url(self):
        u, d = self.v_user.get().strip(), self.v_depot.get().strip()
        self.lbl_url.config(text=("Adresse de la page : " + dl.url_page(u, d)) if u and d else "")

    def rafraichir(self):
        self.lb.delete(0, "end")
        noms = {"apk": "APK", "page": "Page", "autre": "Fichier"}
        for x in self.fichiers:
            nom = x.get("nom") or dl.nom_publie_propose(x["chemin"], x["type"])
            self.lb.insert("end", f"[{noms.get(x['type'], 'Fichier')}]  {nom}      ← {x['chemin']}")

    def ajouter(self, type_):
        types = {"apk": [("Application Android", "*.apk")], "page": [("Page web", "*.html")], "autre": [("Tous les fichiers", "*.*")]}
        ch = filedialog.askopenfilename(parent=self, filetypes=types[type_] + [("Tous", "*.*")])
        if not ch:
            return
        x = {"chemin": ch, "type": type_}
        if type_ == "autre":
            fr = simpledialog.askstring("Nom du fichier", "Nom affiché en FRANÇAIS :", initialvalue=os.path.splitext(os.path.basename(ch))[0], parent=self)
            if not fr:
                return
            mg = simpledialog.askstring("Nom du fichier", "Nom affiché en MALAGASY (vide = même nom) :", parent=self) or ""
            x["libelle_fr"], x["libelle_mg"] = fr.strip(), mg.strip()
        self.fichiers.append(x)
        self.rafraichir()

    def retirer(self):
        for i in sorted(self.lb.curselection(), reverse=True):
            del self.fichiers[i]
        self.rafraichir()

    def parcourir(self):
        d = filedialog.askdirectory(parent=self)
        if d:
            self.v_dossier.set(d)

    def creer(self):
        cfg = self.config()
        pb = dl.verifier(cfg)
        if pb or not cfg["dossier"]:
            messagebox.showerror("Téléchargements", "\n".join(pb) or "Choisissez un dossier.", parent=self)
            return
        try:
            res = dl.construire(cfg, cfg["dossier"], ressource("assets", "telechargement_template.html"),
                                ressource("assets", "logo_banner.png"), ressource("assets", "icone_mobile.png"))
        except ImportError:
            messagebox.showerror("Téléchargements", "Il manque le module « segno ».\nDans PowerShell : py -m pip install segno", parent=self)
            return
        except Exception as e:
            messagebox.showerror("Téléchargements", f"Erreur : {e}", parent=self)
            return
        cfg_app = lire_config()
        cfg_app["telechargements"] = cfg
        ecrire_config(cfg_app)
        if messagebox.askyesno("Téléchargements",
                               "C'est prêt !\n\nAdresse de la page :\n" + res["url"] + "\n\n"
                               "1. Déposez TOUT le contenu du dossier sur votre dépôt GitHub public.\n"
                               "2. Activez GitHub Pages (guide GUIDE_TELECHARGEMENTS_QR.txt).\n"
                               "3. Imprimez Feuille_QR.docx.\n\nOuvrir le dossier maintenant ?", parent=self):
            if hasattr(os, "startfile"):
                try:
                    os.startfile(res["dossier"])
                except OSError:
                    pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("1180x760")
        self.minsize(980, 620)
        self.appliquer_style()
        self.icone()
        self.p = core.nouveau_projet()
        self.chemin = None
        self.modifie = False
        self.cle_selection = None
        self.combos = {}
        self._construire()
        self.charger_vers_ui()
        self.rouvrir_dernier()
        self.protocol("WM_DELETE_WINDOW", self.quitter)

    # ------------------------------------------------------------ apparence
    def icone(self):
        try:
            if sys.platform.startswith("win"):
                self.iconbitmap(ressource("assets", "fdl.ico"))
            else:
                self._img_icone = tk.PhotoImage(file=ressource("assets", "logo_icon.png"))
                self.iconphoto(True, self._img_icone)
        except Exception:
            pass

    def appliquer_style(self):
        self.configure(background=FOND)
        st = ttk.Style(self)
        st.theme_use("clam")
        base = (POLICE, 10)
        gras = (POLICE, 10, "bold")
        st.configure(".", background=FOND, foreground=TEXTE, font=base, bordercolor=BORDURE)
        st.configure("TFrame", background=FOND)
        st.configure("TLabel", background=FOND, foreground=TEXTE)
        st.configure("TLabelframe", background=FOND, bordercolor=BORDURE)
        st.configure("TLabelframe.Label", background=FOND, foreground=NAVY, font=gras)
        st.configure("TCheckbutton", background=FOND)
        st.map("TCheckbutton", background=[("active", FOND)])
        st.configure("TButton", background=NAVY, foreground="white", padding=(10, 5), borderwidth=0,
                     focusthickness=0, font=base)
        st.map("TButton", background=[("pressed", "#0D0452"), ("active", NAVY_CLAIR)],
               foreground=[("disabled", "#AAB")])
        st.configure("TNotebook", background=FOND, borderwidth=0, tabmargins=(0, 4, 0, 0))
        st.configure("TNotebook.Tab", background=FOND_ONGLET, foreground=NAVY, padding=(16, 7), font=gras,
                     borderwidth=0)
        st.map("TNotebook.Tab", background=[("selected", "white")], foreground=[("selected", NAVY)])
        st.configure("Treeview", background="white", fieldbackground="white", rowheight=26, bordercolor=BORDURE)
        st.map("Treeview", background=[("selected", NAVY)], foreground=[("selected", "white")])
        st.configure("Treeview.Heading", background=NAVY, foreground="white", font=gras, relief="flat", padding=5)
        st.map("Treeview.Heading", background=[("active", NAVY_CLAIR)])
        st.configure("TCombobox", fieldbackground="white", background=FOND_ONGLET, arrowcolor=NAVY,
                     selectbackground="white", selectforeground=TEXTE)
        st.map("TCombobox", fieldbackground=[("readonly", "white")], background=[("readonly", FOND_ONGLET)],
               selectbackground=[("readonly", "white")], selectforeground=[("readonly", TEXTE)])
        st.configure("TEntry", fieldbackground="white")
        st.configure("Entete.TFrame", background="white")
        st.configure("EnteteTitre.TLabel", background="white", foreground=NAVY, font=(POLICE, 17, "bold"))
        st.configure("EnteteSous.TLabel", background="white", foreground="#5A5F80", font=(POLICE, 10))

    def entete(self):
        cadre = ttk.Frame(self, style="Entete.TFrame", padding=(14, 8))
        cadre.pack(fill="x")
        try:
            self._img_logo = tk.PhotoImage(file=ressource("assets", "logo_banner.png"))
            ttk.Label(cadre, image=self._img_logo, background="white").pack(side="left", padx=(0, 14))
        except Exception:
            pass
        t = ttk.Frame(cadre, style="Entete.TFrame")
        t.pack(side="left")
        ttk.Label(t, text="FJKM — Programme des dimanches", style="EnteteTitre.TLabel").pack(anchor="w")
        ttk.Label(t, text="Groupes, rôles et export Word du FDL", style="EnteteSous.TLabel").pack(anchor="w")
        tk.Frame(self, height=4, background=NAVY).pack(fill="x")

    # ------------------------------------------------------------ construction
    def _construire(self):
        self.entete()
        barre = ttk.Frame(self, padding=(10, 8))
        barre.pack(fill="x")
        for texte, cmd in (("Nouveau", self.nouveau), ("Ouvrir…", self.ouvrir), ("Enregistrer", self.enregistrer),
                           ("Importer Excel…", self.importer), ("Période suivante", self.suivant)):
            ttk.Button(barre, text=texte, command=cmd).pack(side="left", padx=2)
        barre2 = ttk.Frame(self, padding=(10, 0, 10, 6))
        barre2.pack(fill="x")
        ttk.Label(barre2, text="Exporter :", foreground=NAVY, font=(POLICE, 10, "bold")).pack(side="left", padx=(2, 6))
        for texte, cmd in (("Word…", self.exporter_word), ("Excel…", self.exporter_excel),
                           ("Page mobile…", self.exporter_mobile), ("Publier en ligne…", self.publier_en_ligne),
                           ("Téléchargements (QR)…", self.telechargements)):
            ttk.Button(barre2, text=texte, command=cmd).pack(side="left", padx=2)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.t_reg, self.t_grp, self.t_prog, self.t_eq, self.t_txt = (ttk.Frame(self.nb, padding=10) for _ in range(5))
        self.nb.add(self.t_reg, text="1. Réglages")
        self.nb.add(self.t_grp, text="2. Groupes et membres")
        self.nb.add(self.t_prog, text="3. Programme")
        self.nb.add(self.t_eq, text="4. Équité")
        self.nb.add(self.t_txt, text="5. Langue et textes")
        self._onglet_reglages()
        self._onglet_groupes()
        self._onglet_programme()
        self._onglet_equite()
        self._onglet_textes()
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self.rafraichir_equite())

    def _onglet_reglages(self):
        f = self.t_reg
        f.columnconfigure(1, weight=1)
        self.v_debut, self.v_fin = tk.StringVar(), tk.StringVar()
        self.v_premier = tk.StringVar()
        self.v_tokana, self.v_role5 = tk.BooleanVar(), tk.BooleanVar()

        ttk.Label(f, text="Date du 1er dimanche :").grid(row=0, column=0, sticky="e", pady=4, padx=4)
        ttk.Entry(f, textvariable=self.v_debut, width=14).grid(row=0, column=1, sticky="w")
        ttk.Label(f, text="Date de fin :").grid(row=1, column=0, sticky="e", pady=4, padx=4)
        ttk.Entry(f, textvariable=self.v_fin, width=14).grid(row=1, column=1, sticky="w")
        ttk.Label(f, text="(JJ/MM/AAAA — en général le 30 juin ou le 31 décembre)", foreground="#555").grid(row=1, column=1, padx=130, sticky="w")
        ttk.Label(f, text="Groupe du 1er passage :").grid(row=2, column=0, sticky="e", pady=4, padx=4)
        self.cb_premier = ttk.Combobox(f, textvariable=self.v_premier, state="readonly", width=14)
        self.cb_premier.grid(row=2, column=1, sticky="w")

        ttk.Label(f, text="Ordre de rotation :").grid(row=3, column=0, sticky="ne", pady=8, padx=4)
        o = ttk.Frame(f)
        o.grid(row=3, column=1, sticky="w", pady=8)
        self.lb_ordre = tk.Listbox(o, height=6, width=18, exportselection=False, background="white",
                                   selectbackground=NAVY, selectforeground="white", highlightthickness=1,
                                   highlightbackground=BORDURE, relief="flat", font=(POLICE, 10))
        self.lb_ordre.pack(side="left")
        ob = ttk.Frame(o)
        ob.pack(side="left", padx=6)
        ttk.Button(ob, text="Monter", command=lambda: self.deplacer_ordre(-1)).pack(fill="x", pady=1)
        ttk.Button(ob, text="Descendre", command=lambda: self.deplacer_ordre(1)).pack(fill="x", pady=1)

        ttk.Checkbutton(f, variable=self.v_tokana,
                        text="Dernier dimanche du mois = « Fidirana tokana » (un seul groupe)").grid(row=4, column=1, sticky="w", pady=2)
        ttk.Checkbutton(f, variable=self.v_role5,
                        text="Utiliser aussi le rôle 5 (Vavaka fangatahana) — visible dans l'application et l'Excel, pas dans le Word").grid(row=5, column=1, sticky="w", pady=2)

        ttk.Label(f, text="Adresse du programme en ligne :").grid(row=6, column=0, sticky="e", pady=(10, 0), padx=4)
        u = ttk.Frame(f)
        u.grid(row=6, column=1, sticky="w", pady=(10, 0))
        self.v_url = tk.StringVar()
        ttk.Entry(u, textvariable=self.v_url, width=60).pack(side="left")
        ttk.Button(u, text="Calculer depuis GitHub…", command=self.calculer_url).pack(side="left", padx=6)
        ttk.Label(f, text="Facultatif : permet aux membres de recevoir les mises à jour par internet (voir « Publier en ligne »).",
                  foreground="#555").grid(row=7, column=1, sticky="w")
        b = ttk.Frame(f)
        b.grid(row=8, column=1, sticky="w", pady=14)
        ttk.Button(b, text="Générer le programme et la répartition", command=lambda: self.generer(True)).pack(side="left", padx=2)
        ttk.Button(b, text="Tout régénérer (efface mes corrections)", command=lambda: self.generer(False)).pack(side="left", padx=8)

        aide = ("Comment ça marche : à chaque passage d'un groupe, 4 membres sont choisis — les moins sollicités et les plus "
                "reposés — puis reçoivent chacun le rôle qu'ils ont le moins tenu. Les absents ne sont pas choisis.\n"
                "Vos corrections manuelles (onglet Programme) sont verrouillées : une nouvelle génération les garde et "
                "rééquilibre le reste.")
        ttk.Label(f, text=aide, wraplength=760, justify="left", foreground="#444").grid(row=9, column=0, columnspan=2, sticky="w", pady=6)

    def _onglet_groupes(self):
        f = self.t_grp
        f.columnconfigure(0, weight=1)
        f.rowconfigure(1, weight=1)
        h = ttk.Frame(f)
        h.grid(row=0, column=0, sticky="ew")
        ttk.Label(h, text="Groupe :").pack(side="left")
        self.v_groupe = tk.StringVar()
        self.cb_groupe = ttk.Combobox(h, textvariable=self.v_groupe, state="readonly", width=14)
        self.cb_groupe.pack(side="left", padx=6)
        self.cb_groupe.bind("<<ComboboxSelected>>", lambda e: self.rafraichir_membres())
        ttk.Button(h, text="Nouveau groupe…", command=self.nouveau_groupe).pack(side="left", padx=(20, 2))
        ttk.Button(h, text="Supprimer ce groupe", command=self.supprimer_groupe).pack(side="left", padx=2)
        self.lbl_nb = ttk.Label(h, text="")
        self.lbl_nb.pack(side="right")

        self.tree_m = ttk.Treeview(f, columns=("nom", "abs"), show="headings", selectmode="browse")
        self.tree_m.heading("nom", text="Nom")
        self.tree_m.heading("abs", text="Absences")
        self.tree_m.column("nom", width=260)
        self.tree_m.column("abs", width=520)
        self.tree_m.grid(row=1, column=0, sticky="nsew", pady=8)
        self.tree_m.bind("<Double-1>", lambda e: self.renommer_membre())
        sb = ttk.Scrollbar(f, command=self.tree_m.yview)
        sb.grid(row=1, column=1, sticky="ns", pady=8)
        self.tree_m.config(yscrollcommand=sb.set)

        b = ttk.Frame(f)
        b.grid(row=2, column=0, sticky="w")
        for texte, cmd in (("Ajouter un membre…", self.ajouter_membre), ("Renommer…", self.renommer_membre),
                           ("Supprimer", self.supprimer_membre), ("Absences…", self.absences_membre),
                           ("Monter", lambda: self.deplacer_membre(-1)), ("Descendre", lambda: self.deplacer_membre(1))):
            ttk.Button(b, text=texte, command=cmd).pack(side="left", padx=2)
        ttk.Label(f, text="Deux membres de groupes différents peuvent porter le même nom (ce sont des personnes différentes).",
                  foreground="#555").grid(row=3, column=0, sticky="w", pady=(8, 0))

    def _onglet_programme(self):
        f = self.t_prog
        f.columnconfigure(1, weight=1)
        f.rowconfigure(0, weight=1)
        g = ttk.Frame(f)
        g.grid(row=0, column=0, sticky="ns")
        self.tree_p = ttk.Treeview(g, columns=("date", "g1", "g2", "man"), show="headings", selectmode="browse", height=20)
        for c, t, w in (("date", "Dimanche", 110), ("g1", "Fidirana 1°", 112), ("g2", "Fidirana 2°", 112), ("man", "✎", 32)):
            self.tree_p.heading(c, text=t)
            self.tree_p.column(c, width=w, anchor="center")
        self.tree_p.pack(side="left", fill="y")
        sb = ttk.Scrollbar(g, command=self.tree_p.yview)
        sb.pack(side="left", fill="y")
        self.tree_p.config(yscrollcommand=sb.set)
        self.tree_p.bind("<<TreeviewSelect>>", self.selection_dimanche)

        d = ttk.Frame(f)
        d.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        d.columnconfigure(0, weight=1)
        d.rowconfigure(1, weight=1)
        self.lbl_dim = ttk.Label(d, text="Choisissez un dimanche à gauche.", font=("TkDefaultFont", 11, "bold"))
        self.lbl_dim.grid(row=0, column=0, sticky="w")
        self.detail = ttk.Frame(d)
        self.detail.grid(row=1, column=0, sticky="nsew", pady=6)
        ttk.Label(d, text="À vérifier :").grid(row=2, column=0, sticky="w")
        self.txt_av = tk.Text(d, height=7, wrap="word", state="disabled", background="#FFF8E1", relief="flat",
                              highlightthickness=1, highlightbackground=BORDURE, font=(POLICE, 10))
        self.txt_av.grid(row=3, column=0, sticky="ew")

    def _onglet_equite(self):
        f = self.t_eq
        f.columnconfigure(0, weight=1)
        f.rowconfigure(1, weight=1)
        h = ttk.Frame(f)
        h.grid(row=0, column=0, sticky="ew")
        ttk.Label(h, text="Groupe :").pack(side="left")
        self.v_eq = tk.StringVar()
        self.cb_eq = ttk.Combobox(h, textvariable=self.v_eq, state="readonly", width=14)
        self.cb_eq.pack(side="left", padx=6)
        self.cb_eq.bind("<<ComboboxSelected>>", lambda e: self.rafraichir_equite())
        self.lbl_eq = ttk.Label(h, text="")
        self.lbl_eq.pack(side="left", padx=16)
        cols = ("nom", "per", "r1", "r2", "r3", "r4", "r5", "tot", "der")
        self.tree_e = ttk.Treeview(f, columns=cols, show="headings")
        for c, t, w in (("nom", "Nom", 200), ("per", "Participations (période)", 150), ("r1", "VVK VKT", 70), ("r2", "VKT 1", 60),
                        ("r3", "VKT 2", 60), ("r4", "VKT 3", 60), ("r5", "Rôle 5", 60),
                        ("tot", "Total avec historique", 150), ("der", "Dernier passage", 110)):
            self.tree_e.heading(c, text=t)
            self.tree_e.column(c, width=w, anchor="w" if c == "nom" else "center")
        self.tree_e.grid(row=1, column=0, sticky="nsew", pady=8)
        ttk.Label(f, text="Participations : combien de fois chaque membre est choisi sur la période. "
                          "L'écart entre le plus et le moins sollicité reste d'une participation quand personne n'est absent.",
                  foreground="#555", wraplength=900, justify="left").grid(row=2, column=0, sticky="w")

    def _onglet_textes(self):
        f = self.t_txt
        f.columnconfigure(0, weight=1)
        f.rowconfigure(2, weight=1)
        h = ttk.Frame(f)
        h.grid(row=0, column=0, sticky="w")
        ttk.Label(h, text="Langue proposée par défaut :").pack(side="left")
        self.v_langue = tk.StringVar(value="Français")
        cb = ttk.Combobox(h, textvariable=self.v_langue, values=["Français", "Malagasy"], state="readonly", width=12)
        cb.pack(side="left", padx=8)
        cb.bind("<<ComboboxSelected>>", lambda e: self.changer_langue())
        self.v_langue_libre = tk.BooleanVar(value=False)
        ttk.Checkbutton(h, text="Laisser les membres changer de langue (bouton Malagasy / Français)", variable=self.v_langue_libre,
                        command=self.changer_langue_libre).pack(side="left", padx=16)
        ttk.Label(f, text="Si la case est décochée, tous les membres voient la langue choisie ici (pas de bouton de langue). "
                          "Si elle est cochée, chaque membre peut changer lui-même de langue ; son choix est gardé sur son téléphone, "
                          "même après une mise à jour en ligne. Vos textes personnalisés ne s'appliquent qu'à la langue choisie ici. "
                          "La traduction malagasy est un brouillon à relire. Tout est envoyé avec « Page mobile… » ou « Publier en ligne… ».",
                  wraplength=1000, justify="left").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.tree_t = ttk.Treeview(f, columns=("el", "fr", "mg", "perso"), show="headings", selectmode="browse")
        for c, t, w in (("el", "Élément", 250), ("fr", "Français", 270), ("mg", "Malagasy (brouillon)", 270), ("perso", "Votre texte", 270)):
            self.tree_t.heading(c, text=t)
            self.tree_t.column(c, width=w)
        self.tree_t.grid(row=2, column=0, sticky="nsew", pady=8)
        sb = ttk.Scrollbar(f, command=self.tree_t.yview)
        sb.grid(row=2, column=1, sticky="ns", pady=8)
        self.tree_t.config(yscrollcommand=sb.set)
        self.tree_t.bind("<Double-1>", lambda e: self.modifier_texte())
        b = ttk.Frame(f)
        b.grid(row=3, column=0, sticky="w")
        ttk.Button(b, text="Modifier…", command=self.modifier_texte).pack(side="left", padx=2)
        ttk.Button(b, text="Remettre le texte de la langue", command=self.effacer_texte).pack(side="left", padx=2)
        ttk.Label(f, text="{g} = nom du groupe, {n} et {m} = nombres, {s}… selon l'élément.",
                  foreground="#555").grid(row=4, column=0, sticky="w", pady=(6, 0))

    def changer_langue_libre(self):
        self.p["langue_libre"] = bool(self.v_langue_libre.get())
        self.marquer()

    def changer_langue(self):
        self.p["langue"] = "mg" if self.v_langue.get() == "Malagasy" else "fr"
        self.rafraichir_textes()
        self.marquer()

    def rafraichir_textes(self):
        self.v_langue.set("Malagasy" if self.p.get("langue") == "mg" else "Français")
        self.v_langue_libre.set(bool(self.p.get("langue_libre", False)))
        self.tree_t.delete(*self.tree_t.get_children())
        perso = self.p.get("textes", {})
        for cle, desc, fr in core.TEXTES_PAGE:
            self.tree_t.insert("", "end", iid=cle, values=(desc, fr, core.TEXTES_MG.get(cle, ""), perso.get(cle, "")))

    def modifier_texte(self):
        sel = self.tree_t.selection()
        if not sel:
            return
        cle = sel[0]
        desc, fr = next((d, f) for k, d, f in core.TEXTES_PAGE if k == cle)
        mg = core.TEXTES_MG.get(cle, "")
        base = mg if self.p.get("langue") == "mg" and mg else fr          # texte de la langue choisie
        actuel = self.p.get("textes", {}).get(cle, "")
        v = simpledialog.askstring("Texte de la page", f"{desc}\n\nFrançais : {fr}\nMalagasy : {mg}",
                                   initialvalue=actuel or base, parent=self)
        if v is None:
            return
        v = v.strip()
        textes = self.p.setdefault("textes", {})
        if not v or v == base:
            textes.pop(cle, None)
        else:
            textes[cle] = v
        self.rafraichir_textes()
        self.marquer()

    def effacer_texte(self):
        sel = self.tree_t.selection()
        if sel and self.p.get("textes", {}).pop(sel[0], None) is not None:
            self.rafraichir_textes()
            self.marquer()

    # ------------------------------------------------------------ état <-> interface
    def titre(self):
        nom = os.path.basename(self.chemin) if self.chemin else "sans titre"
        self.title(f"FJKM · FDL — {nom}{' *' if self.modifie else ''}")

    def marquer(self):
        self.modifie = True
        self.titre()

    def charger_vers_ui(self):
        p = self.p
        self.v_debut.set(fmt(p["debut"]))
        self.v_fin.set(fmt(p["fin"]))
        self.v_tokana.set(p["tokana"])
        self.v_role5.set(p["role5"])
        self.v_url.set(p.get("url_maj", ""))
        self.rafraichir_ordre()
        self.v_premier.set(p["premier_groupe"])
        self.cb_groupe["values"] = list(p["ordre"])
        self.cb_eq["values"] = list(p["ordre"])
        if self.v_groupe.get() not in p["ordre"]:
            self.v_groupe.set(p["ordre"][0] if p["ordre"] else "")
        if self.v_eq.get() not in p["ordre"]:
            self.v_eq.set(p["ordre"][0] if p["ordre"] else "")
        self.rafraichir_membres()
        self.rafraichir_programme()
        self.rafraichir_equite()
        self.rafraichir_textes()
        self.rafraichir_avert([])
        self.titre()

    def lire_reglages(self):
        try:
            debut, fin = core.lire_date(self.v_debut.get()), core.lire_date(self.v_fin.get())
        except Exception:
            messagebox.showerror("Date", "Dates invalides. Format attendu : JJ/MM/AAAA.")
            return False
        if fin < debut:
            messagebox.showerror("Date", "La date de fin est avant la date de début.")
            return False
        if debut.weekday() != 6:
            messagebox.showinfo("Date", f"Le {fmt(debut)} est un {JOURS[debut.weekday()]} : "
                                        f"je prends le dimanche suivant ({fmt(core.prochain_dimanche(debut))}).")
            debut = core.prochain_dimanche(debut)
            self.v_debut.set(fmt(debut))
        p = self.p
        p["debut"], p["fin"] = core.d2s(debut), core.d2s(fin)
        p["premier_groupe"] = self.v_premier.get() or (p["ordre"][0] if p["ordre"] else "")
        p["tokana"], p["role5"] = self.v_tokana.get(), self.v_role5.get()
        url = self.v_url.get().strip()
        if url and not url.startswith("https://"):
            messagebox.showerror("Adresse", "L'adresse du programme en ligne doit commencer par https://")
            return False
        p["url_maj"] = url
        if url:
            cfg = lire_config()
            if cfg.get("url_maj") != url:
                cfg["url_maj"] = url
                ecrire_config(cfg)
        return True

    # ------------------------------------------------------------ réglages
    def rafraichir_ordre(self):
        self.lb_ordre.delete(0, "end")
        for g in self.p["ordre"]:
            self.lb_ordre.insert("end", g)
        self.cb_premier["values"] = list(self.p["ordre"])

    def deplacer_ordre(self, delta):
        sel = self.lb_ordre.curselection()
        if not sel:
            return
        i, j = sel[0], sel[0] + delta
        o = self.p["ordre"]
        if 0 <= j < len(o):
            o[i], o[j] = o[j], o[i]
            self.rafraichir_ordre()
            self.lb_ordre.selection_set(j)
            self.marquer()

    def generer(self, garder_verrous):
        if not self.lire_reglages():
            return
        if not self.p["ordre"]:
            messagebox.showerror("Groupes", "Aucun groupe.")
            return
        if not garder_verrous and any(self.p["verrous"]) and not messagebox.askyesno(
                "Tout régénérer", "Vos corrections manuelles seront effacées. Continuer ?"):
            return
        try:
            core.generer_programme(self.p)
            avert = core.generer_affectations(self.p, garder_verrous)
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
            return
        self.v_debut.set(fmt(self.p["debut"]))
        self.rafraichir_programme()
        self.rafraichir_equite()
        self.rafraichir_avert(avert + [v for v in core.verifier(self.p) if v not in avert])
        self.marquer()
        self.nb.select(self.t_prog)

    # ------------------------------------------------------------ groupes et membres
    def groupe(self):
        return self.v_groupe.get()

    def membre_selectionne(self):
        sel = self.tree_m.selection()
        return self.tree_m.item(sel[0], "values")[0] if sel else None

    @staticmethod
    def resume_absences(abs_):
        return " ; ".join(f"du {fmt(a)} au {fmt(b)}" for a, b in abs_)

    def rafraichir_membres(self):
        self.tree_m.delete(*self.tree_m.get_children())
        g = self.groupe()
        if g not in self.p["groupes"]:
            return
        for m in self.p["groupes"][g]["membres"]:
            self.tree_m.insert("", "end", iid=m["nom"], values=(m["nom"], self.resume_absences(m.get("absences", []))))
        self.lbl_nb.config(text=f"{len(self.p['groupes'][g]['membres'])} membres")

    def ajouter_membre(self):
        g = self.groupe()
        if not g:
            return
        nom = simpledialog.askstring("Nouveau membre", f"Nom du membre du groupe {g} :", parent=self)
        if nom:
            try:
                core.ajouter_membre(self.p, g, nom)
            except ValueError as e:
                messagebox.showerror("Membre", str(e))
                return
            self.rafraichir_membres()
            self.marquer()

    def renommer_membre(self):
        g, ancien = self.groupe(), self.membre_selectionne()
        if not ancien:
            return
        nouveau = simpledialog.askstring("Renommer", "Nouveau nom :", initialvalue=ancien, parent=self)
        if nouveau and nouveau.strip() != ancien:
            try:
                core.renommer_membre(self.p, g, ancien, nouveau)
            except ValueError as e:
                messagebox.showerror("Membre", str(e))
                return
            self.rafraichir_membres()
            self.rafraichir_programme()
            self.marquer()

    def supprimer_membre(self):
        g, nom = self.groupe(), self.membre_selectionne()
        if not nom:
            return
        if messagebox.askyesno("Supprimer", f"Retirer « {nom} » du groupe {g} ?\n\n"
                                            "Les passages déjà générés où il/elle figure restent inchangés "
                                            "jusqu'à la prochaine génération."):
            core.supprimer_membre(self.p, g, nom)
            self.rafraichir_membres()
            self.marquer()

    def absences_membre(self):
        g, nom = self.groupe(), self.membre_selectionne()
        if not nom:
            return
        for m in self.p["groupes"][g]["membres"]:
            if m["nom"] == nom:
                dlg = DialogAbsences(self, f"Absences — {nom}", m.get("absences", []))
                m["absences"] = dlg.absences
        self.rafraichir_membres()
        self.marquer()

    def deplacer_membre(self, delta):
        g, nom = self.groupe(), self.membre_selectionne()
        if not nom:
            return
        ms = self.p["groupes"][g]["membres"]
        i = next(k for k, m in enumerate(ms) if m["nom"] == nom)
        j = i + delta
        if 0 <= j < len(ms):
            ms[i], ms[j] = ms[j], ms[i]
            self.rafraichir_membres()
            self.tree_m.selection_set(nom)
            self.marquer()

    def nouveau_groupe(self):
        nom = simpledialog.askstring("Nouveau groupe", "Nom du groupe :", parent=self)
        if not nom:
            return
        nom = nom.strip().upper()
        if nom in self.p["groupes"]:
            messagebox.showerror("Groupe", "Ce groupe existe déjà.")
            return
        self.p["groupes"][nom] = {"membres": []}
        self.p["ordre"].append(nom)
        self.v_groupe.set(nom)
        self.charger_vers_ui()
        self.marquer()

    def supprimer_groupe(self):
        g = self.groupe()
        if not g or len(self.p["ordre"]) <= 2:
            messagebox.showinfo("Groupe", "Il faut garder au moins deux groupes.")
            return
        if messagebox.askyesno("Supprimer", f"Supprimer le groupe {g} et tous ses membres ?"):
            del self.p["groupes"][g]
            self.p["ordre"].remove(g)
            if self.p["premier_groupe"] == g:
                self.p["premier_groupe"] = self.p["ordre"][0]
            self.v_groupe.set(self.p["ordre"][0])
            self.charger_vers_ui()
            self.marquer()

    # ------------------------------------------------------------ programme
    def nb_roles(self):
        return 5 if self.p["role5"] else 4

    def rafraichir_programme(self):
        sel = self.cle_selection
        self.tree_p.delete(*self.tree_p.get_children())
        verrous = set(self.p["verrous"])
        for it in self.p["programme"]:
            gs = it["groupes"]
            manuel = "✎" if any(f"{it['date']}|{g}" in verrous for g in gs) else ""
            self.tree_p.insert("", "end", iid=it["date"],
                               values=(fmt(it["date"]), gs[0], gs[1] if len(gs) > 1 else "(tokana)", manuel))
        if sel and self.tree_p.exists(sel):
            self.tree_p.selection_set(sel)
        else:
            self.vider_detail()

    def vider_detail(self):
        for w in self.detail.winfo_children():
            w.destroy()
        self.combos = {}
        self.cle_selection = None
        self.lbl_dim.config(text="Choisissez un dimanche à gauche." if self.p["programme"]
                            else "Allez dans « Réglages » puis cliquez sur « Générer ».")

    def selection_dimanche(self, _=None):
        sel = self.tree_p.selection()
        if not sel:
            return
        date_iso = sel[0]
        self.cle_selection = date_iso
        item = next(it for it in self.p["programme"] if it["date"] == date_iso)
        for w in self.detail.winfo_children():
            w.destroy()
        self.combos = {}
        d = core.s2d(date_iso)
        tokana = "  —  FIDIRANA TOKANA" if len(item["groupes"]) == 1 else ""
        self.lbl_dim.config(text=f"Dimanche {fmt(d)}{tokana}")
        for col, g in enumerate(item["groupes"]):
            cadre = ttk.LabelFrame(self.detail, text=f" {g} ", padding=8)
            cadre.grid(row=0, column=col, sticky="n", padx=6, pady=2)
            cle = f"{date_iso}|{g}"
            aff = self.p["affectations"].get(cle, {})
            noms = [""] + core.noms_du_groupe(self.p, g)
            for r in range(1, self.nb_roles() + 1):
                ttk.Label(cadre, text=f"{ROLES_COURTS[r]} :", width=9).grid(row=r, column=0, sticky="e", pady=2)
                v = tk.StringVar(value=aff.get(str(r), ""))
                cb = ttk.Combobox(cadre, textvariable=v, values=noms, state="readonly", width=22)
                cb.grid(row=r, column=1, pady=2, padx=4)
                cb.bind("<<ComboboxSelected>>", lambda e, c=cle, rr=r, vv=v: self.changer_role(c, rr, vv.get()))
                self.combos[(cle, r)] = v
            verrouille = cle in self.p["verrous"]
            ttk.Label(cadre, text="✎ modifié à la main (conservé)" if verrouille else "calculé automatiquement",
                      foreground="#a15c00" if verrouille else "#777").grid(row=9, column=0, columnspan=2, pady=(6, 0))
            if verrouille:
                ttk.Button(cadre, text="Redonner la main à l'application",
                           command=lambda c=cle: self.deverrouiller(c)).grid(row=10, column=0, columnspan=2, pady=4)

    def changer_role(self, cle, role, nom):
        aff = self.p["affectations"].setdefault(cle, {str(r): "" for r in range(1, self.nb_roles() + 1)})
        aff[str(role)] = nom
        if cle not in self.p["verrous"]:
            self.p["verrous"].append(cle)
        self.marquer()
        self.maj_marques()
        self.rafraichir_equite()
        self.rafraichir_avert(core.verifier(self.p))
        self.after(20, self.selection_dimanche)

    def maj_marques(self):
        verrous = set(self.p["verrous"])
        for it in self.p["programme"]:
            manuel = "✎" if any(f"{it['date']}|{g}" in verrous for g in it["groupes"]) else ""
            if self.tree_p.exists(it["date"]):
                self.tree_p.set(it["date"], "man", manuel)

    def deverrouiller(self, cle):
        if cle in self.p["verrous"]:
            self.p["verrous"].remove(cle)
        avert = core.generer_affectations(self.p, True)
        self.maj_marques()
        self.rafraichir_equite()
        self.rafraichir_avert(avert + [v for v in core.verifier(self.p) if v not in avert])
        self.marquer()
        self.after(20, self.selection_dimanche)

    def rafraichir_avert(self, lignes):
        self.txt_av.config(state="normal")
        self.txt_av.delete("1.0", "end")
        self.txt_av.insert("end", "\n".join(lignes) if lignes else "Rien à signaler.")
        self.txt_av.config(state="disabled")

    # ------------------------------------------------------------ équité
    def rafraichir_equite(self):
        self.tree_e.delete(*self.tree_e.get_children())
        g = self.v_eq.get()
        if not self.p["programme"] or g not in self.p["groupes"]:
            self.lbl_eq.config(text="")
            return
        lignes = core.statistiques(self.p).get(g, [])
        for l in lignes:
            self.tree_e.insert("", "end", values=(l["nom"], l["periode"], *(l["roles"][r] for r in (1, 2, 3, 4, 5)),
                                                  l["total"], fmt(l["dernier"]) if l["dernier"] else "—"))
        if lignes:
            per = [l["periode"] for l in lignes]
            self.lbl_eq.config(text=f"Moins sollicité : {min(per)} fois — plus sollicité : {max(per)} fois")

    # ------------------------------------------------------------ fichiers
    def confirmer_perte(self):
        if self.modifie:
            r = messagebox.askyesnocancel("FDL", "Enregistrer les modifications ?")
            if r is None:
                return False
            if r and not self.enregistrer():
                return False
        return True

    def rouvrir_dernier(self):
        """Au démarrage : rouvre automatiquement le dernier fichier enregistré ou ouvert."""
        ch = lire_config().get("dernier")
        if ch and os.path.exists(ch):
            try:
                self.remplacer_projet(core.charger(ch), ch)
                if self.p["programme"]:
                    self.nb.select(self.t_prog)
            except Exception:
                pass

    def remplacer_projet(self, p, chemin=None):
        self.p, self.chemin, self.modifie, self.cle_selection = p, chemin, False, None
        if not self.p.get("url_maj") and lire_config().get("url_maj"):
            self.p["url_maj"] = lire_config()["url_maj"]       # même adresse en ligne pour tous vos projets
        if chemin:
            cfg = lire_config()
            cfg["dernier"] = chemin
            ecrire_config(cfg)
        self.charger_vers_ui()

    def nouveau(self):
        if self.confirmer_perte():
            self.remplacer_projet(core.nouveau_projet())

    def ouvrir(self):
        if not self.confirmer_perte():
            return
        ch = filedialog.askopenfilename(filetypes=[("Projet FDL", "*.json"), ("Tous", "*.*")])
        if ch:
            try:
                self.remplacer_projet(core.charger(ch), ch)
            except Exception as e:
                messagebox.showerror("Ouvrir", f"Fichier illisible :\n{e}")

    def enregistrer(self):
        if not self.lire_reglages():
            return False
        if not self.chemin:
            ch = filedialog.asksaveasfilename(defaultextension=".json", initialfile="programme_FDL.json",
                                              filetypes=[("Projet FDL", "*.json")])
            if not ch:
                return False
            self.chemin = ch
        try:
            core.sauver(self.p, self.chemin)
        except OSError as e:
            messagebox.showerror("Enregistrer", str(e))
            return False
        self.modifie = False
        self.titre()
        cfg = lire_config()
        cfg["dernier"] = self.chemin
        ecrire_config(cfg)
        return True

    def importer(self):
        if not self.confirmer_perte():
            return
        ch = filedialog.askopenfilename(title="Classeur Excel existant",
                                        filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tous", "*.*")])
        if not ch:
            return
        avec = messagebox.askyesno(
            "Importer",
            "Importer aussi le programme déjà rempli (les noms de chaque dimanche) ?\n\n"
            "OUI : je garde les noms tels quels (marqués « modifié à la main »).\n"
            "NON : je ne garde que les groupes, les membres et les réglages ; "
            "la répartition est recalculée par l'application.")
        try:
            p = core.importer_excel(ch, avec)
            if not avec:
                core.generer_programme(p)
                core.generer_affectations(p)
        except Exception as e:
            messagebox.showerror("Importer", f"Import impossible :\n{e}")
            return
        self.remplacer_projet(p)
        self.modifie = True
        self.titre()
        self.rafraichir_avert(core.verifier(p))
        self.nb.select(self.t_prog if p["programme"] else self.t_reg)

    def exporter_word(self):
        if not self.lire_reglages():
            return
        if not self.p["programme"]:
            messagebox.showinfo("Word", "Générez d'abord le programme (onglet Réglages).")
            return
        manque = [v for v in core.verifier(self.p) if "non attribué" in v]
        if manque and not messagebox.askyesno("Word", f"{len(manque)} rôle(s) ne sont pas attribués "
                                                       "(cases vides dans le Word). Continuer quand même ?"):
            return
        d0 = core.s2d(self.p["debut"])
        dlg = DialogSemestre(self, "S1" if d0.month <= 6 else "S2", d0.year)
        if not dlg.resultat:
            return
        sem, annee = dlg.resultat
        ch = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=f"Programme_FDL_{sem}_{annee}.docx",
                                          filetypes=[("Word", "*.docx")])
        if not ch:
            return
        try:
            manquants = ew.creer_word(core.vers_dimanches(self.p), annee, sem, ch)
        except PermissionError:
            messagebox.showerror("Word", "Impossible d'écrire le fichier : fermez-le dans Word, puis réessayez.")
            return
        except Exception as e:
            messagebox.showerror("Word", f"Erreur : {e}")
            return
        if manquants:
            messagebox.showwarning("Word", "Fichier créé, mais ces dimanches du semestre ne sont pas dans le programme :\n"
                                           + ", ".join(manquants[:12]) + (" …" if len(manquants) > 12 else ""))
        else:
            messagebox.showinfo("Word", "Le fichier Word est créé.")
        if hasattr(os, "startfile"):
            try:
                os.startfile(ch)
            except OSError:
                pass

    def exporter_excel(self):
        if not self.p["programme"]:
            messagebox.showinfo("Excel", "Générez d'abord le programme.")
            return
        ch = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Programme_FDL.xlsx",
                                          filetypes=[("Excel", "*.xlsx")])
        if not ch:
            return
        try:
            core.exporter_excel(self.p, ch)
            messagebox.showinfo("Excel", "Fichier Excel créé.")
        except Exception as e:
            messagebox.showerror("Excel", str(e))

    def exporter_mobile(self):
        if not self.p["programme"]:
            messagebox.showinfo("Page mobile", "Générez d'abord le programme.")
            return
        manque = [v for v in core.verifier(self.p) if "non attribué" in v]
        if manque and not messagebox.askyesno("Page mobile", f"{len(manque)} rôle(s) ne sont pas attribués. "
                                                              "Continuer quand même ?"):
            return
        ch = filedialog.asksaveasfilename(defaultextension=".html", initialfile="FDL_Mon_Planning.html",
                                          filetypes=[("Page web", "*.html")])
        if not ch:
            return
        try:
            core.exporter_mobile(self.p, ch, ressource("assets", "mobile_template.html"),
                                 ressource("assets", "logo_banner.png"), ressource("assets", "icone_mobile.png"))
        except Exception as e:
            messagebox.showerror("Page mobile", f"Erreur : {e}")
            return
        messagebox.showinfo(
            "Page mobile",
            "La page est créée.\n\n"
            "Envoyez ce fichier aux membres (WhatsApp, Messenger, e-mail…) ou publiez-le sur un site : "
            "chacun l'ouvre dans Chrome, choisit son groupe et son nom une seule fois, puis consulte son "
            "planning et ajoute ses rappels à son calendrier.\n\n"
            "Avec l'application Android FDL, le membre touche « Charger le programme » et choisit ce fichier.\n"
            "La page est en lecture seule : un membre ne peut rien modifier.\n"
            "À chaque nouveau programme, renvoyez la nouvelle page — ou utilisez « Publier en ligne » pour "
            "que les membres le reçoivent tout seuls.")

    def calculer_url(self):
        user = simpledialog.askstring("GitHub", "Votre nom d'utilisateur GitHub :", initialvalue=GITHUB_UTILISATEUR, parent=self)
        if not user:
            return
        depot = simpledialog.askstring("GitHub", "Nom du dépôt PUBLIC où déposer le programme :",
                                       initialvalue="fdl-programme", parent=self)
        if not depot:
            return
        self.v_url.set(f"https://raw.githubusercontent.com/{user.strip()}/{depot.strip()}/main/programme.json")
        self.marquer()

    def publier_en_ligne(self):
        if not self.lire_reglages():
            return
        if not self.p["programme"]:
            messagebox.showinfo("Publier en ligne", "Générez d'abord le programme.")
            return
        url = self.p.get("url_maj", "")
        if not url:
            messagebox.showinfo("Publier en ligne",
                                "Renseignez d'abord l'adresse du programme en ligne :\n"
                                "onglet « Réglages », bouton « Calculer depuis GitHub… ».")
            self.nb.select(self.t_reg)
            return
        manque = [v for v in core.verifier(self.p) if "non attribué" in v]
        if manque and not messagebox.askyesno("Publier en ligne", f"{len(manque)} rôle(s) ne sont pas attribués. "
                                                                   "Continuer quand même ?"):
            return
        nom = url.split("?")[0].rstrip("/").split("/")[-1] or "programme.json"
        ch = filedialog.asksaveasfilename(defaultextension=".json", initialfile=nom,
                                          filetypes=[("Programme en ligne", "*.json")])
        if not ch:
            return
        if os.path.basename(ch) != nom and not messagebox.askyesno(
                "Nom du fichier", f"L'adresse en ligne se termine par « {nom} » mais vous enregistrez « "
                                  f"{os.path.basename(ch)} ».\nLes membres ne le recevraient pas. Continuer quand même ?"):
            return
        try:
            core.exporter_programme_json(self.p, ch)
        except Exception as e:
            messagebox.showerror("Publier en ligne", f"Erreur : {e}")
            return
        texte = ("Le fichier est prêt.\n\nPour que les membres le reçoivent :\n"
                 "1. ouvrez la page GitHub de votre dépôt public ;\n"
                 "2. glissez ce fichier dessus (il remplace l'ancien) ;\n"
                 "3. cliquez sur « Commit changes ».\n\n"
                 "Les membres le reçoivent à leur prochaine ouverture de l'application, avec internet "
                 "(comptez 5 minutes au plus après la publication).")
        m = re.match(r"https://raw\.githubusercontent\.com/([^/]+)/([^/]+)/([^/]+)/", url)
        if m and messagebox.askyesno("Publier en ligne", texte + "\n\nOuvrir la page GitHub maintenant ?"):
            webbrowser.open(f"https://github.com/{m.group(1)}/{m.group(2)}/upload/{m.group(3)}")
        elif not m:
            messagebox.showinfo("Publier en ligne", texte)
        self.marquer()

    def telechargements(self):
        DialogTelechargements(self, lire_config().get("telechargements", {}))

    def suivant(self):
        if not self.p["programme"]:
            messagebox.showinfo("Période suivante", "Générez d'abord le programme de la période actuelle.")
            return
        if not messagebox.askyesno(
                "Période suivante",
                "Créer la période suivante à partir de celle-ci ?\n\n"
                "Les groupes, membres et réglages sont conservés ; l'historique des participations est repris "
                "pour que la suite reste équitable. Pensez à enregistrer la période actuelle d'abord."):
            return
        if not self.lire_reglages():
            return
        try:
            q = core.continuer(self.p)
            core.generer_programme(q)
            avert = core.generer_affectations(q)
        except Exception as e:
            messagebox.showerror("Période suivante", str(e))
            return
        self.remplacer_projet(q)
        self.modifie = True
        self.titre()
        self.rafraichir_avert(avert + [v for v in core.verifier(q) if v not in avert])
        self.nb.select(self.t_prog)

    def quitter(self):
        if self.confirmer_perte():
            self.destroy()


def autotest():
    """py fdl_app.py --selftest : vérifie que tout est en place (génération + Word) sans ouvrir de fenêtre."""
    import tempfile
    p = core.nouveau_projet(dt.date(2026, 7, 5))
    p["tokana"] = True
    for g in p["ordre"]:
        for i in range(8):
            core.ajouter_membre(p, g, f"Membre {g} {i + 1}")
    core.generer_programme(p)
    avert = core.generer_affectations(p)
    sortie = os.path.join(tempfile.gettempdir(), "fdl_autotest.docx")
    manquants = ew.creer_word(core.vers_dimanches(p), 2026, "S2", sortie)
    ok = os.path.exists(sortie) and not manquants and not avert
    for f in ("fdl.ico", "logo_banner.png", "logo_icon.png", "icone_mobile.png", "mobile_template.html"):
        ok = ok and os.path.exists(ressource("assets", f))
    page = os.path.join(tempfile.gettempdir(), "fdl_autotest_mobile.html")
    core.exporter_mobile(p, page, ressource("assets", "mobile_template.html"),
                         ressource("assets", "logo_banner.png"), ressource("assets", "icone_mobile.png"))
    with open(page, encoding="utf-8") as fh:
        ok = ok and "Mon planning" in fh.read()
    dossier_dl = os.path.join(tempfile.gettempdir(), "fdl_autotest_dl")
    fic = os.path.join(tempfile.gettempdir(), "fdl_autotest.docx")
    res = dl.construire({"utilisateur": "test", "depot": "essai", "fichiers": [{"chemin": fic, "type": "autre"}]}, dossier_dl,
                        ressource("assets", "telechargement_template.html"), ressource("assets", "logo_banner.png"),
                        ressource("assets", "icone_mobile.png"))
    ok = ok and all(os.path.exists(os.path.join(dossier_dl, n)) for n in ("index.html", "QR_telechargements.png", "Feuille_QR.docx"))
    pj = os.path.join(tempfile.gettempdir(), "fdl_autotest_programme.json")
    core.exporter_programme_json(p, pj)
    with open(pj, encoding="utf-8") as fh:
        ok = ok and "dimanches" in json.load(fh)
    print("AUTOTEST", "OK" if ok else "ECHEC", sortie)
    return 0 if ok else 1


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        sys.exit(autotest())
    App().mainloop()
