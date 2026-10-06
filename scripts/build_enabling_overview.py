"""Build the "Enabling Pulse – Management Overview" slide.

Same layout as the use case overview (build_management_overview.py), with rows
that fit the enabling streams: a short goal and one key KPI instead of phase and
E2E go-live date, and four dimensions (Use Cases instead of Data Provisioning /
Supplier Activation).

    python scripts/build_enabling_overview.py [template.pptx] [output.pptx]
"""

import sys

from build_management_overview import BG, CW, TEMPLATE, USE_CASE_BOARD, build, compute_layout, write_pptx

TEMPLATE = sys.argv[1] if len(sys.argv) > 1 else TEMPLATE
OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "output/Enabling_Pulse_Management_Overview_CW41_26.pptx"

ENABLING_DIMS = ["External · CX Association", "Business", "Use Cases", "Technical"]

# CW 41 / 2026 Enabling Pulse slides. Content in short business English, abbreviations lower case.
# ms = next critical milestone (red before amber, then earliest; Portfolio: Secure PR74 as agreed);
# None = no amber/red milestone. kpi = (KPI title, actual, target) from the "KPIs (defined per area)" box;
# actual None = not reported in the source slide (shown as "tbd").
COLUMNS = [  # sorted: Escalation needed -> Problem solving -> On track
    dict(name="Supplier Activation", abbr="sac-cx", dl="Stuhrmann (BZ-PX)", status="esc",
         goal="All P-Mat suppliers operationally connected",
         kpi=("Suppliers contracted by Q3", "240", "280"),
         ms=("esc", "280 suppliers contracted", "30 Sep 2026"),
         dims=["esc", "ps", "ns", "ns"],
         decisions=[(1, "How do we make Cofinity-X onboarding clear for suppliers, and who covers "
                        "the data sharing agreement (DSA) costs per use case?")]),
    dict(name="Portfolio, AI+, C-X NEXT, DigiTrace", abbr="port-cx", dl="Kraus (K-DDX/5)", status="esc",
         goal="Budget and funding secured in every planning round",
         kpi=("Budgets secured", "0", "2 (PR74, PR75)"),
         ms=("esc", "Secure PR74", "Nov 2026"),
         dims=["ok", "ps", "ps", "ok"],
         decisions=[(2, "PR75: which costs are planned per use case?"),
                    (3, "New PR75 process: capacity crunch & high risks?"),
                    (4, "DigiTrace: stop, or fund for 2027?")]),
    dict(name="Data Provisioning", abbr="dpr-cx", dl="Kraus (K-DDX/5)", status="ps",
         goal="Catena-X data provided as shoppable data products",
         kpi=("Use Case Owners approved", "6", "8"), ms=None,
         dims=["ok", "ok", "ps", "ns"]),
    dict(name="Hub & Spoke", abbr="hub-cx", dl="Kraus (K-DDX/5)", status="ps",
         goal="Catena-X scaled across all brands via Hub & Spoke",
         kpi=("Brands signed", "8", "17"),
         ms=("ps", "Framework agreements signed", "31 Dec 2026"),
         dims=["ok", "ps", "ns", "ns"]),
    dict(name="Legal & Contracts", abbr="gis-cx", dl="Brunke (K-ILX-6), Schiffelbaum (A-IIED-EJ)", status="ps",
         goal="Framework contracts with all brands and data consumers",
         kpi=("Automated contract cycle", "not live", "live by 01 Jan 2028"),
         ms=("ps", "Managed by Cofinity-X", "Oct 2026"),
         dims=["ok", "ps", "ok", "ps"]),
    dict(name="Architecture", abbr="arch-cx", dl="Fischer (K-DCC/B)", status="ps",
         goal="Approved, scalable Catena-X target architecture",
         kpi=None, ms=None,
         dims=["ok", "ok", "ok", "ps"]),
]

ENABLING_BOARD = dict(
    eyebrow="CATENA-X",
    title="Management Overview of Enabling Streams",
    head_labels=["Name", "Workstream Lead"],
    columns=COLUMNS,
    dims=ENABLING_DIMS,
    dim_gap=0.08,  # four dimensions: more air between them
    rows=[
        ("status", "Overall status"),
        ("goal", "Goal"),
        ("kpi", ("Key KPI", "actual / target")),
        ("ms", "Next critical milestone"),
        ("dims", "Dimensions"),
        ("dec", "Decision required"),
    ],
    notes=f"Enabling Pulse Management Overview {CW} – one card per enabling stream (read top to bottom), "
          "sorted by overall status. Next critical milestone = most critical amber/red milestone.",
)


def main():
    layout = compute_layout([USE_CASE_BOARD, ENABLING_BOARD])  # same grid as the use case slide
    write_pptx(build(ENABLING_BOARD, layout).shapes, OUTPUT, ENABLING_BOARD["notes"], bg=BG)


if __name__ == "__main__":
    main()
