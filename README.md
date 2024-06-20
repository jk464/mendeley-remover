# mendeley-remover
Replace references using Mendeley Reference Manager with plain text

References in word doc broken cause MRM is rubbish? This allows you to replace those references with plaintext ones (assuming you've already made).

Install requirements with `python pip install python-docx`

Then run `python mrm_remover.py <DOCX>`

This will run through the Word Doc (trying it's best to replace) the MRM references with plaintext ones.

If there's any it can't fix it'll output the reference it can't fix.

If any references are still wrong you can add a mapping in `reference_map` to map the incorrect reference to the correct one, i.e.

```
reference_map = {
  "INCORRECT REFERENCE 1": "CORRECT REFERENCE 1",
  "INCORRECT REFERENCE 2": "CORRECT REFERENCE 2",
}
```

No promises this will work - make a backup of your docx first - this script was hacked together to fix my parents thesis in the last days of her PhD.

Yes this script is stupidly bad - yes it runs in `O(N^2)` instead of `O(N)` time - the removing of references breaks the for loop and I couldn't be _fucked_ trying to fix it properly cause it worked.
