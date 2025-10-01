import io
import re
import pandas as pd
import streamlit as st


st.set_page_config(page_title="Payroll Consolidation", layout="centered")
st.title("Payroll Consolidation Tool")

st.markdown("""
**Instructions**
1) Upload the **Reference Columns** Excel workbook.  
2) Upload all **payroll files** for this run (Excel .xlsx).  
3) Check the file list and total.  
4) Confirm correctness, then click **Process files**.  
5) Download the consolidated Excel.
""")

# ---------- Inputs ----------
ref_file = st.file_uploader(
    "Upload Reference Columns (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=False
)
payroll_files = st.file_uploader(
    "Upload Payroll Files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
) or []

st.subheader("Uploaded Files")
st.write("Reference file:")
st.write("- " + ref_file.name if ref_file else "- None")

st.write("Payroll files:")
if payroll_files:
    for uf in payroll_files:
        st.write(f"- {uf.name}")
else:
    st.write("- None")

st.write(f"Total payroll files: {len(payroll_files)}")

st.subheader("Confirm and Process")
confirmed = st.checkbox("I confirm the files are complete and correct.")
process_btn = st.button("Process files")

# ---------- Core helpers (minimal) ----------
def process_payroll_file(file_like) -> pd.DataFrame:
    """Standardize one payroll file."""
    df = pd.read_excel(file_like, header=3, engine="openpyxl")

    # drop first two blank rows if present
    if len(df.index) >= 2:
        df = df.drop(df.index[:2])

    # keep rows with valid Employee ID pattern 000-000000
    if "Employee ID*" in df.columns:
        pattern = r"^\d{3}-\d{6}$"
        df = df[df["Employee ID*"].astype(str).str.match(pattern, na=False)]

    # get filename (without extension) from upload
    name = getattr(file_like, "name", "")
    period = name.rsplit(".", 1)[0] if name else ""

    # split "PayrollType-MMddyyyy"
    df["Payroll Period"] = period
    split_df = df["Payroll Period"].str.split("-", n=1, expand=True)
    if split_df.shape[1] < 2:
        split_df = split_df.reindex(columns=[0, 1])
    df["Payroll Type"] = split_df[0]
    df["Payroll Date"] = split_df[1]

    # derive date/cutoff/month
    df["Payroll Date"] = pd.to_datetime(df["Payroll Date"], format="%m%d%Y", errors="coerce")
    df["Cutoff Type"] = df["Payroll Date"].dt.day.map(
        lambda x: "First Cutoff" if x == 15 else "Second Cutoff" if x in [30, 31] else "Other"
    )
    df["Month"] = df["Payroll Date"].dt.strftime("%B")

    df = df.drop(columns=["Payroll Period"], errors="ignore")

    leading = ["Payroll Type", "Payroll Date", "Month", "Cutoff Type"]
    remaining = [c for c in df.columns if c not in leading]
    return df[leading + remaining]


def derive_ref_columns(ref_uploaded) -> list:
    """Build reference column order by processing the reference file once."""
    # Make a fresh copy because read_excel consumes the buffer
    b = io.BytesIO(ref_uploaded.getvalue())
    b.name = ref_uploaded.name
    ref_df = process_payroll_file(b)
    return list(ref_df.columns)


def reorder_like_reference(final_df: pd.DataFrame, ref_cols: list) -> pd.DataFrame:
    """Reorder columns to match reference; append extras near similar names if possible."""
    all_cols = list(final_df.columns)
    new_order = list(ref_cols)

    extras = [c for c in all_cols if c not in ref_cols]
    for col in extras:
        base = re.sub(r"\d+$", "", col)
        near = [i for i, c in enumerate(new_order) if base and base in c]
        if near:
            new_order.insert(near[-1] + 1, col)
        else:
            new_order.append(col)

    # ensure uniqueness
    out, seen = [], set()
    for c in new_order:
        if c not in seen and c in all_cols:
            out.append(c)
            seen.add(c)

    # include any missing columns
    for c in all_cols:
        if c not in seen:
            out.append(c)
            seen.add(c)

    return final_df.reindex(columns=out)


def consolidate(payroll_uploads, ref_cols):
    """Concat processed files, skip duplicate base filenames."""
    dfs, processed, skipped = [], [], []
    seen_bases = set()

    for uf in payroll_uploads:
        name = uf.name
        if not name.lower().endswith(".xlsx"):
            continue
        base = name.rsplit(".", 1)[0]
        if base in seen_bases:
            skipped.append(name)
            continue
        seen_bases.add(base)

        b = io.BytesIO(uf.getvalue())
        b.name = uf.name
        try:
            df = process_payroll_file(b)
            dfs.append(df)
            processed.append(name)
        except Exception as e:
            st.warning(f"Error processing {name}: {e}")

    if not dfs:
        return pd.DataFrame(), processed, skipped

    final_df = pd.concat(dfs, ignore_index=True)
    final_df = reorder_like_reference(final_df, ref_cols)
    return final_df, processed, skipped


# ---------- Run ----------
if process_btn:
    if not ref_file:
        st.error("Please upload the Reference Columns workbook.")
    elif len(payroll_files) == 0:
        st.error("Please upload at least one payroll file.")
    elif not confirmed:
        st.error("Please confirm that the files are complete and correct.")
    else:
        try:
            ref_columns = derive_ref_columns(ref_file)
        except Exception as e:
            st.error(f"Unable to derive reference columns: {e}")
            st.stop()

        final_df, processed_files, skipped_dups = consolidate(payroll_files, ref_columns)

        if final_df.empty:
            st.warning("No files were processed.")
        else:
            st.success("Consolidation completed.")
            st.write(f"Files processed: {len(processed_files)}")
            for n in processed_files:
                st.write(f"- {n}")
            if skipped_dups:
                st.write(f"Skipped duplicate base filenames: {len(skipped_dups)}")
                for n in skipped_dups:
                    st.write(f"- {n}")

            # write Excel in memory using openpyxl (no extra writers)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                final_df.to_excel(writer, index=False, sheet_name="Consolidated")
            buffer.seek(0)

            st.download_button(
                "Download consolidated Excel",
                data=buffer,
                file_name="Payroll_Consolidation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
