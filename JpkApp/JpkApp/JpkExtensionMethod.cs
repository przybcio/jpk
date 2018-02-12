using JpkApp.Model;

namespace JpkApp
{
    public static class JpkExtensionMethod
    {
        public static void InitZakupArray(this JPK jpk, int zakupLenght)
        {
            jpk.ZakupWiersz = new JPKZakupWiersz[zakupLenght];
        }

        public static void InitSprzedazArray(this JPK jpk, int sprzedazLenght)
        {
            jpk.SprzedazWiersz = new JPKSprzedazWiersz[sprzedazLenght];
        }
    }
}