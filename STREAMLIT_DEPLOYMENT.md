# Deploy the z/OS learning lab to Streamlit

This project adds the **MVS learning lab** to the existing product compatibility
dashboard. The 3270-style console is an educational simulation: it has no host
connection and does not submit JCL or change a z/OS system.

## Files

Keep these files and the existing `data/` directory together:

- `app.py`
- `zos_learning_lab.py`
- `compatibility_db.py`
- `rag_engine.py`
- `vendor_details.py`
- `requirements.txt`
- `data/`

`requirements.txt` is the filename Streamlit Community Cloud expects (not
`requirment.txt`). It includes `joblib`; the lab can download a serialized
TF-IDF retrieval index as `zos_learning_rag_index.joblib`.

## Deploy

1. Push the project to a private or public GitHub repository approved for your
   data and environment.
2. In Streamlit Community Cloud, choose **Create app**, select the repository
   and branch, and set the main file path to `app.py`.
3. Deploy. Streamlit installs the dependencies from `requirements.txt`, including
   `openpyxl` for the Excel exports. After changing dependencies, push the
   updated `requirements.txt` to the selected branch and reboot or redeploy the
   app so Streamlit rebuilds its environment.
4. Open **MVS learning lab** and try `HELP`, `LISTCAT`, `ISPF`, `SDSF`, `D A`,
   and `SUBMIT` with the sample JCL.

The compatibility dashboard retrieves vendor information at startup. The host
running the app therefore needs approved outbound HTTPS access to its existing
vendor source links.

## Optional IBM watsonx.ai answers

Without credentials, the tutor still retrieves notes from its local TF-IDF
index and displays the related IBM documentation references. To enable optional
Granite-generated answers, add secrets in the Streamlit app's **Settings >
Secrets** area:

```toml
WATSONX_APIKEY = "your-IBM-Cloud-API-key"
WATSONX_PROJECT_ID = "your-watsonx-project-id"
WATSONX_URL = "https://us-south.ml.cloud.ibm.com"
WATSONX_MODEL_ID = "ibm/granite-3-3-8b-instruct"
```

Use either `WATSONX_PROJECT_ID` or `WATSONX_SPACE_ID`, not both. Choose a URL
and model enabled for your IBM Cloud account. Never put secrets in source
control. If model access fails, the tutor shows an error and the locally
retrieved notes remain available.

## Logo and certificate

After passing the learning assessment (4/5 or better), a learner can enter a
name and optionally upload a PNG/JPG logo to create a printable HTML certificate.
The certificate is marked **unofficial** and explicitly states that it is not
issued, endorsed, or accredited by IBM. Upload only artwork you are authorized
to use. The app does not ship an IBM logo.

## IBM documentation

Topic lessons link to IBM Documentation for z/OS 3.2.0 and identify the relevant
IBM manuals in each lesson. Documentation may require navigating/searching by
the displayed manual title. Confirm commands and procedures against the release
and site configuration before using them on a real system.
