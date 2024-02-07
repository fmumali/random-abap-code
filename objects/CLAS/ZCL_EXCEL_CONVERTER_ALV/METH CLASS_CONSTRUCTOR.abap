  METHOD class_constructor.
* let's fill the color conversion routines.
    DATA: ls_color TYPE ts_col_converter.
* 0 all combination the same
    ls_color-col       = 0.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 0.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 0.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 0.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

* Blue
    ls_color-col       = 1.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFB0E4FC'. " 176 228 252 blue
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 1.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFB0E4FC'. " 176 228 252 blue
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 1.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FF5FCBFE'. " 095 203 254 Int blue
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 1.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FF5FCBFE'. " 095 203 254
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255
    INSERT ls_color INTO TABLE wt_colors.

* Gray
    ls_color-col       = 2.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'.
    ls_color-fillcolor = 'FFE5EAF0'. " 229 234 240 gray
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 2.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFE5EAF0'. " 229 234 240 gray
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 2.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFD8E8F4'. " 216 234 244 int gray
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 2.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFD8E8F4'. " 216 234 244 int gray
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

*Yellow
    ls_color-col       = 3.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFEFEB8'. " 254 254 184 yellow
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 3.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFFEFEB8'. " 254 254 184 yellow
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 3.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFF9ED5D'. " 249 237 093 int yellow
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 3.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFF9ED5D'. " 249 237 093 int yellow
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

* light blue
    ls_color-col       = 4.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFCEE7FB'. " 206 231 251 light blue
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 4.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFCEE7FB'. " 206 231 251 light blue
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 4.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FF9ACCEF'. " 154 204 239 int light blue
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 4.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FF9ACCEF'. " 154 204 239 int light blue
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

* Green
    ls_color-col       = 5.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFCEF8AE'. " 206 248 174 Green
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 5.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFCEF8AE'. " 206 248 174 Green
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 5.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FF7AC769'. " 122 199 105 int Green
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 5.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FF7AC769'. " 122 199 105 int Green
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

* Red
    ls_color-col       = 6.
    ls_color-int       = 0.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFDBBBC'. " 253 187 188 Red
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 6.
    ls_color-int       = 0.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFFDBBBC'. " 253 187 188 Red
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 6.
    ls_color-int       = 1.
    ls_color-inv       = 0.
    ls_color-fontcolor = 'FF000000'. " 000 000 000 Black
    ls_color-fillcolor = 'FFFB6B6B'. " 251 107 107 int Red
    INSERT ls_color INTO TABLE wt_colors.

    ls_color-col       = 6.
    ls_color-int       = 1.
    ls_color-inv       = 1.
    ls_color-fontcolor = 'FFFB6B6B'. " 251 107 107 int Red
    ls_color-fillcolor = 'FFFFFFFF'. " 255 255 255 White
    INSERT ls_color INTO TABLE wt_colors.

  ENDMETHOD.