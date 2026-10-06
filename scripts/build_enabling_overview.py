"""Build the "Enabling Pulse – Management Overview" slide.

Same layout as the use case overview (build_management_overview.py), with rows
that fit the enabling streams: a short goal and one key KPI instead of phase and
E2E go-live date, and four dimensions (Use Cases instead of Data Provisioning /
Supplier Activation).

    python scripts/build_enabling_overview.py [template.pptx] [output.pptx]
"""

import sys

from build_management_overview import BG, CW, TEMPLATE, build, write_pptx

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else TEMPLATE
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "output/Enabling_Pulse_Management_Overview_CW41_26.pptx"

ENABLING_DIMS = ["External · CX Association", "Business", "Use Cases", "Technical"]

# CW 41 / 2026 Enabling Pulse slides. Content in short business English, abbreviations lower case.
# ms = next critical milestone (red before amber, then earliest; Portfolio: Secure PR74 as agreed);
# None = no amber/red milestone. kpi = (KPI title, value) taken from the "KPIs (defined per area)" box.
COLUMNS = [  # sorted: Escalation needed -> Problem solving -> On track
    dict(name="Supplier Activation", abbr="sac-cx", dl="Behrens (BZ-PX)", status="esc",
         goal="All P-Mat suppliers operationally connected",
         kpi=("Contracts by end of Q3 2026", "240 of 280 suppliers"),
         sup=(240, 400), sup_pct=55, sup_note="228 company groups",
         ms=("esc", "280 suppliers contracted", "30 Sep 2026"),
         dims=["esc", "ps", "ns", "ns"]),
    dict(name="Portfolio, AI+, C-X NEXT, DigiTrace", abbr="port-cx", dl="Bollmann, Poetsch (K-DDX/5)", status="esc",
         goal="Budget and funding secured in every planning round",
         kpi=("PR74 benefit / costs", "~650 T€ / ~340 T€"),
         sup=None, sup_text="–",
         ms=("esc", "Secure PR74", "Nov 2026"),
         dims=["ok", "ps", "ps", "ok"],
         decisions=[(1, "PR75: which costs are planned per use case?"),
                    (2, "New PR75 process: capacity crunch & high risks?"),
                    (3, "DigiTrace: stop, or fund for 2027?")]),
    dict(name="Data Provisioning", abbr="dpr-cx", dl="Bollmann, Poetsch (K-DDX/5)", status="ps",
         goal="CX data provided operationally as shoppable products",
         kpi=("Use Case Owners", "6 named & approved"),
         sup=None, sup_text="–", ms=None,
         dims=["ok", "ok", "ps", "ns"]),
    dict(name="Hub & Spoke", abbr="hub-cx", dl="Bollmann, Poetsch (K-DDX/5)", status="ps",
         goal="Catena-X scaled across all brands via Hub & Spoke",
         kpi=("Brands connected", "8 of 17 signed"),
         sup=None, sup_text="–",
         ms=("ps", "Framework agreements signed", "31 Dec 2026"),
         dims=["ok", "ps", "ns", "ns"]),
    dict(name="Legal & Contracts", abbr="gis-cx", dl="Pietschmann", status="ps",
         goal="Framework contracts for brands and data consumers",
         kpi=("Automated contract cycle", "Target 01 Jan 2028"),
         sup=None, sup_text="–",
         ms=("ps", "Managed by Cofinity-X", "Oct 2026"),
         dims=["ok", "ps", "ok", "ps"]),
    dict(name="Architecture", abbr="arch-cx", dl="Fischer", status="ps",
         goal="Approved, scalable Catena-X target architecture",
         kpi=None,
         sup=None, sup_text="–", ms=None,
         dims=["ok", "ok", "ok", "ps"]),
]

ENABLING_BOARD = dict(
    eyebrow="ENABLING PULSE · CATENA-X",
    title="Management Overview of Enabling Streams",
    columns=COLUMNS,
    dims=ENABLING_DIMS,
    dim_gap=0.08,  # four dimensions: more air between them
    rows=[
        ("status", "Overall status"),
        ("goal", "Goal"),
        ("kpi", "Key KPI"),
        ("sup", "Suppliers enabled 2026"),
        ("ms", "Next critical milestone"),
        ("dims", "Dimensions"),
        ("dec", "Decision required"),
    ],
    notes=f"Enabling Pulse Management Overview {CW} – one card per enabling stream (read top to bottom), "
          "sorted by overall status. Next critical milestone = most critical amber/red milestone.",
)


def main():
    write_pptx(build(ENABLING_BOARD).shapes, OUTPUT, ENABLING_BOARD["notes"], bg=BG)


if __name__ == "__main__":
    main()
