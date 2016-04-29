using Microsoft.SharePoint;
using Microsoft.SharePoint.Publishing.Navigation;
using Microsoft.SharePoint.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StructurePublishWebSite
{
    class Program
    {
        static string sphost = "devsharepoint"; // SharePoint Host

        // 3.	Crear una función para definir los termsets
        static void CreateNavigationTermSet(Group group, string name)
        {
            TermSet termSet = null;

            try
            {
                termSet = group.TermSets[name];
            }
            catch (Exception ex) { }

            if (termSet == null)
            {
                //create term set
                termSet = group.CreateTermSet(name);
                termSet.Description = name;
                termSet.IsAvailableForTagging = true;
                termSet.IsOpenForTermCreation = true;

                //set properties
                termSet.SetCustomProperty("_Sys_Nav_IsNavigationTermSet", "True");
                termSet.TermStore.CommitAll();
            }
        }

        // 4.	Crear una función para definir los términos
        static void CreateNavigationTerms(TermSet termSet)
        {
            //create terms
            Term t = termSet.CreateTerm("Home", 1033);
            t.SetLocalCustomProperty("_Sys_Nav_TargetUrl", "/Pages/default.aspx");

            t = termSet.CreateTerm("Intranet", 1033);
            t.SetLocalCustomProperty("_Sys_Nav_SimpleLinkUrl", "http://"+ sphost +"/sites/publicacion/Pages/Home.aspx");

            t = termSet.CreateTerm("HR", 1033);
            t.SetLocalCustomProperty("_Sys_Nav_SimpleLinkUrl", "http://"+ sphost + "/sites/publicacion/hr");

            t = termSet.CreateTerm("IT", 1033);
            t.SetLocalCustomProperty("_Sys_Nav_SimpleLinkUrl", "http://" + sphost + "/sites/publicacion/it");

            t = termSet.CreateTerm("Legal", 1033);
            t.SetLocalCustomProperty("_Sys_Nav_SimpleLinkUrl", "http://" + sphost + "/sites/publicacion/legal");

            t = termSet.CreateTerm("Finance", 1033);
            t.SetLocalCustomProperty("_Sys_Nav_SimpleLinkUrl", "http://" + sphost + "/sites/publicacion/finance");

            try
            {
                termSet.TermStore.CommitAll();
            }
            catch (Exception ex) { }
        }

        // 5.	Crear una función para definir los grupos
        static Group CreateGroup(TermStore termStore, string name)
        {
            Group g = null;

            try
            {
                g = termStore.Groups[name];
            }
            catch (Exception ex) { }

            if (g == null)
            {
                g = termStore.CreateGroup(name);
                termStore.CommitAll();
            }

            return g;
        }

        // 6 Crear una función para fijar la navegación
        static void PinTermSet(TermSet source, TermSet target)
        {
            foreach (Term t in source.Terms)
            {
                target.ReuseTermWithPinning(t);
            }
         }

        // 7.	Crear el código en la función Main que cree los grupos, términos y los enlace para la navegación

        static void Main(string[] args)
        {
            SPSecurity.RunWithElevatedPrivileges(
                () =>
                {
                    using (SPSite site = new SPSite("http://" + sphost + "/sites/publicacion"))
                    {
                        TaxonomySession session = new TaxonomySession(site);
                        TermStore termStore = session.TermStores["Managed Metadata Service"];
                        Group group = CreateGroup(termStore, "Navigation");
                        CreateNavigationTermSet(group, "Intranet");

                        TermSet s_termSet = group.TermSets["Intranet"];
                        CreateNavigationTerms(s_termSet);

                        CreateNavigationTermSet(group, "Team");
                        TermSet t_termSet = group.TermSets["Team"];

                        PinTermSet(s_termSet, t_termSet);
                        termStore.CommitAll();
                    }



                    using (SPSite site = new SPSite("http://" + sphost + "/sites/publicacion"))

                    {

                        SetManagedNavigation(site, "Navigation", "Intranet");

                    }

                });

        }

        static void SetManagedNavigation(SPSite site, string groupName, string termSetName)

        {

            TaxonomySession session = new TaxonomySession(site);

            TermStore termStore = session.TermStores["Managed Metadata Service"];

            Group group = termStore.Groups[groupName];

            TermSet termSet = group.TermSets[termSetName];

            WebNavigationSettings settings = new WebNavigationSettings(site.RootWeb);

            settings.GlobalNavigation.Source = StandardNavigationSource.TaxonomyProvider;

            settings.GlobalNavigation.TermStoreId = termStore.Id;

            settings.GlobalNavigation.TermSetId = termSet.Id;

            settings.Update();

        }
    }
}
