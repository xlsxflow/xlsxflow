# Security

## Reporting a vulnerability

Please report security issues privately, not in a public issue. Email palikaomkar.22.cse@anits.edu.in with a description and, if you can, a file or script that shows the problem. You will get a reply within 7 days.

Fixes go into the latest release. Older versions are not patched.

## Scope

Reports we want, for `@xlsxflow/core` and `@xlsxflow/pro`:

- a crafted `.xlsx`, `.xlsm` or CSV file that crashes the reader, hangs it, or makes it use far more memory than the `maxUncompressedBytes` limit allows;
- output that lets data passed to the writer or editor change the structure of the file (for example, inject XML or formulas);
- a way to make `@xlsxflow/pro` accept a licence key that was not issued for a paid order.

Values are returned as the file stores them. A hyperlink that is a `javascript:` URL, or text that starts with `=`, is not a vulnerability in the library: check such values before putting them in a web page or a CSV file. See "Reading files from untrusted users" in the README.
