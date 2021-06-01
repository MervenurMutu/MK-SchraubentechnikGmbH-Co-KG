namespace Schraubenprogramm
{
    internal class Schraube
    {
        public double laenge { get; set; }
        public double gewindeLaenge { get; set; }
        public string gewindeart { get; set; }
        public string gewinde { get; set; }
        public double k { get; set; }
        public double innenradius { get; set; }
        public double steigung { get; set; }

        public Schraube(double laenge, double gewindeLaenge, string gewindeart,
            string gewinde, double k, double innenradius, double steigung)
        {
            this.laenge = laenge;
            this.gewindeLaenge = gewindeLaenge;
            this.gewindeart = gewindeart;
            this.gewinde = gewinde;
            this.k = k;
            this.innenradius = innenradius;
            this.steigung = steigung;
        }
    }
    



}





