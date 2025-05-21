# MS Excel things

Showcase of some of the MS Excel things I've previously done.

## Attribute Collator

![Screenshot of first sheet in "Attribute Collator.xlsx"](https://github.com/jsjs2401/ms-excel-things/blob/main/images/Attribute%20Collator.png)

Collates attributes from input data where each entry corresponds to one attribute for a specific item, with the output presenting all available attributes in one row for each item.

Also removes empty and duplicate attributes.

### Equivalent code in Python:

```
input = [["water", "wet"], ["water", "slippery"], ["fire", "hot"], ["water", "refreshing"], ["water", "refreshing"],
         ["fire", "panas"], ["wood", "hard"], ["fire", "atsui"], ["metal", "harder"], ["wood", ""], ["metal", "heavy"]]


def attrCollator(input):
    output = dict()
    for x in input:
        if x[1] == "":
            continue
        if x[0] in output.keys():
            if x[1] not in output[x[0]]:
                output[x[0]].append(x[1])
        else:
            output[x[0]] = [x[1]]
    return output
```

## Organization Hierarchy Formatter

![Screenshot of first sheet in "Organization Hierarchy Formatter.xlsx"](https://github.com/jsjs2401/ms-excel-things/blob/main/images/Organization%20Hierarchy%20Formatter.png)

Converts relational organizational data into a chart form.

Includes some error-checking such as:

- Checking for duplicate organizations.
- Checking for potentially missing relational links.
- Some capacity to identify corrupted input data.

### Equivalent code in Python:

```
input = [["HQ", ""], ["Office 1", "HQ"], ["Office 2", "HQ"], ["Office 3", "HQ"], ["Office 4", "HQ"], ["Office 5", "HQ"],
         ["Office 6", "HQ"], ["Suboffice 1-1", "Office 1"], ["Suboffice 1-2", "Office 1"], ["Suboffice 1-3", "Office 1"],
         ["Suboffice 2-1", "Office 2"], ["Suboffice 3-1", "Office 3"], ["Suboffice 3-2", "Office 3"],
         ["Suboffice 4-1", "Office 4"], ["Suboffice 5-1", "Office 5"], ["Suboffice 6-1", "Office 6"]]


def orgHierarchyFormatter(input):
    organizations = dict()
    output = dict()
    for x in input:
        if x[0] in organizations.keys():
            print(f"Warning: Duplicated organization {x}")
        else:
            organizations[x[0]] = x[1]
    for org in organizations.keys():
        if org not in organizations.values() and org not in output.keys():
            output[org] = deque([org])
    for x in output.keys():
        while output[x][0] != "":
            try:
                output[x].appendleft(organizations[output[x][0]])
            except KeyError:
                print(f"Warning: Missing relational link for {output[x][0]}")
                break
    return output
```
