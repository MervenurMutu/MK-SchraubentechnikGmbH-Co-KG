internal class Innensechskantkopf
{
    public double zylinderdurchmesser { get; set; }
    public double höhe { get; set; }
    public double innenhöhe { get; set; }
    public double innenschlüsselweite { get; set; }
    public string art { get; set; }

    public Innensechskantkopf(double zylinderdurchmesser, double höhe, double innehöhe, double innenschlüsselweite, string art)
    {
        this.zylinderdurchmesser = zylinderdurchmesser;
        this.höhe = höhe;
        this.innenhöhe = innehöhe;
        this.innenschlüsselweite = innenschlüsselweite;
        this.art = "Innensechskant";
    }
}
    









