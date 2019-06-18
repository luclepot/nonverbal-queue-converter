import argparse
import sys

def group_by_color(
    filepath,
    row_height,
    column_width,
    include_sources,
    output_name,
):
    from StyleFrame import StyleFrame, Styler, utils
    import numpy as np
    import pandas as pd
   
    if not filepath.endswith(".xlsx"):
        filepath += ".xlsx"
        
    sf = StyleFrame.read_excel(filepath, read_style=True)
    color_groupings = {}

    for c_index, c in enumerate(sf.columns):
        if "Cues" in c.value:
            for value in sf[c]:
                if type(value.value) is not str and np.isnan(value.value):
                    continue
                if value.style.bg_color + "_value" not in color_groupings:
                    color_groupings[value.style.bg_color + "_value"] = []
                if include_sources and (
                    value.style.bg_color + "_source" not in color_groupings
                ):
                    color_groupings[value.style.bg_color + "_source"] = []
                color_groupings[value.style.bg_color + "_value"].append(
                    value.value
                )
                if include_sources:
                    color_groupings[value.style.bg_color + "_source"].append(
                        sf.iloc[0,c_index - 1].value
                    )

    # colors of our stuff
    for k,length in zip(color_groupings.keys(), map(len, color_groupings.values())):
        print( "color", k + ":", length, "entries")

    # pad data so it is all the same size
    maxlen = max(map(len, color_groupings.values()))
    for k,v in color_groupings.items():
        color_groupings[k] = v + ['']*(maxlen - len(v))
        print("old length:", len(v), "|| new length:", len(color_groupings[k]))

    # add padded data to df, print 'er out
    new_df = pd.DataFrame(color_groupings)
    new_df

    defaults = {'font': utils.fonts.calibri, 'font_size': 9}

    new_sf = StyleFrame(new_df, styler_obj=Styler(**defaults))

    for c in new_sf.columns:

        # color to use
        color = c.value.strip("_source").strip("_value")

        # apply style to ALL values (incl headers)
        new_sf.apply_column_style(
            cols_to_style=[c.value],
            styler_obj=Styler(bold=True, font_size=10, bg_color=color,
            font_color=utils.colors.white),
            style_header=True,
        )

        # revert style for non-headers
        new_sf.apply_column_style(
            cols_to_style=[c.value],
            styler_obj=Styler(**defaults),
            style_header=False
        )

    # row height
    all_rows = new_sf.row_indexes
    new_sf.set_row_height_dict(
        row_height_dict={
            all_rows[0]: 24,
            all_rows[1:]: row_height
        }
    )

    # col width
    all_cols = tuple(map(lambda x: x.value, new_sf.columns))
    new_sf.set_column_width_dict(
        col_width_dict={
            all_cols: column_width
        }
    )

    # save to excel file
    if not output_name.endswith(".xlsx"):
        output_name += ".xlsx"
    new_sf.to_excel(output_name).save()
    
    return 0

if __name__ == "__main__":
    parser = argparse.ArgumentParser("group by color")
    parser.add_argument("input")
    parser.add_argument("-o", "--output", required=False, default="grouped", dest="output", help="defaults to 'grouped'")
    parser.add_argument("-s", "--include-sources", dest="sources", action="store_true", help="defaults to False")
    parser.add_argument("-r", "--row-height", type=float, dest="rowh", default=21, required=False, help="defaults to 21")
    parser.add_argument("-c", "--column-width", type=float, dest="colw", default=18, required=False, help="defaults to 18")
    
    args = parser.parse_args(sys.argv[1:])
    group_by_color(
        args.input,
        args.rowh,
        args.colw,
        args.sources,
        args.output
    )
    
