





# Documasonry

Template Engine for Word / Excel AutoCAD

- jinja2 template
- yaml config
- detect name and doc body fields
- generate vertex table







filler



word_filler

excel_filler

autocad_filler

  text object

  block object basepoint on center basepoint on original

  explode block object



### usage

``` python
filler = Filler.from_template(path='')
filler.detect_required_fields()

filler.render(info=yaml_info)

filler.save(folder='path/to/', close=True)



info = Information.from_yaml(path='yaml_path')


```