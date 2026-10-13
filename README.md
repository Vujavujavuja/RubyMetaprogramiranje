# Ruby Metaprogramming: Google Sheets as Ruby objects

A small Ruby library that wraps a Google Sheets worksheet and lets you work with it like a native Ruby table. Homework for the Scripting Languages course, written to practise metaprogramming.

## What it does

```ruby
t = Sheet.new(worksheet)

t.to_a                  # the table as a two-dimensional array
t.row(1)                # one row, indexed like an array
t.each { |cell| ... }   # Enumerable over every cell, left to right

t["Prva kolona"]        # a column by its header
t["Prva kolona"][1]     # a cell in that column
t.prvaKolona            # the same column through a generated method
t.prvaKolona.sum        # column aggregates: sum, avg
t.prvaKolona.map { |c| c.to_i + 1 }
t.index.rn1422          # the row whose "index" cell is rn1422

t1 + t2                 # append the rows of two tables with matching headers
t1 - t2                 # remove the rows of t2 from t1
```

It also skips merged cells and ignores any row containing the keywords `total` or `subtotal`.

The column methods are not written by hand: they are defined at runtime from the header row with `define_singleton_method`, and row lookups by cell value go through `method_missing`.

## Running it

```bash
gem install google_drive
ruby domaci1.rb
```

The script expects a `config.json` with Google OAuth credentials next to it, in the format the [google_drive](https://github.com/gimite/google-drive-ruby) gem uses. That file is ignored by git: never commit it.
