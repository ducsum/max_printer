import os
import glob

with open('/data/sum_printer/filter_printer/source_code.md', 'w', encoding='utf-8') as out:
    for f in sorted(glob.glob('/data/sum_printer/filter_printer/*.py')):
        basename = os.path.basename(f)
        out.write(f'[{basename}]:\n')
        with open(f, 'r', encoding='utf-8') as inf:
            out.write(inf.read())
        out.write('\n\n')
