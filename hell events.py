from functions import *

construction_gear = 381.9
VIP_LEVEL = 13  # These numbers depend on your account, those are mine

for construction in constructions:
    wb.active = constructions.index(construction)  # Select the sheet that corresponds to every construction
    find_titles()  # Find out where all the important cells are (for more info, check the function in functions.py)
    might_table.append(assign_mights(constructions[constructions.index(construction)]))
    time_table.append(assign_times(constructions[constructions.index(construction)]))
    food_table.append(assign_foods(constructions[constructions.index(construction)]))
    stone_table.append(assign_stones(constructions[constructions.index(construction)]))
    timber_table.append(assign_timbers(constructions[constructions.index(construction)]))
    ore_table.append(assign_ores(constructions[constructions.index(construction)]))

i_index = 0

for i in time_table:  # Translate every item in time_table to seconds
    x_index = 0
    for x in i:
        time_table[i_index][x_index] = translate_time(time_table[i_index][x_index])
        x_index += 1
    i_index += 1


i_index = 0

for i in might_table:
    x_index = 0
    for x in i:
        might_dict[f'{constructions[i_index]}_{x_index + 1}'] = might_table[i_index][x_index]
        time_dict[f'{constructions[i_index]}_{x_index + 1}'] = time_table[i_index][x_index]
        food_dict[f'{constructions[i_index]}_{x_index + 1}'] = food_table[i_index][x_index]
        stone_dict[f'{constructions[i_index]}_{x_index + 1}'] = stone_table[i_index][x_index]
        timber_dict[f'{constructions[i_index]}_{x_index + 1}'] = timber_table[i_index][x_index]
        ore_dict[f'{constructions[i_index]}_{x_index + 1}'] = ore_table[i_index][x_index]
        rss_dict[f'{constructions[i_index]}_{x_index + 1}'] = 0
        try:  # Some values aren't numbers (N/A, None, etc), so it just sums up as many as it can for every construction
            rss_dict[f'{constructions[i_index]}_{x_index + 1}'] += food_table[i_index][x_index]
        except TypeError:
            pass
        try:
            rss_dict[f'{constructions[i_index]}_{x_index + 1}'] += stone_table[i_index][x_index]
        except TypeError:
            pass
        try:
            rss_dict[f'{constructions[i_index]}_{x_index + 1}'] += timber_table[i_index][x_index]
        except TypeError:
            pass
        try:
            rss_dict[f'{constructions[i_index]}_{x_index + 1}'] += ore_table[i_index][x_index]
        except TypeError:
            pass



        try:  # Some values are "", None, N/A, etc (for example, castle is already lv1 when you start the game)
            ratio_conversor[f'{constructions[i_index]}_{x_index + 1}'] = int(might_table[i_index][x_index]) / int(translate_time(time_table[i_index][x_index]))
        except (ValueError, AttributeError, ZeroDivisionError, TypeError):
            ratio_conversor[f'{constructions[i_index]}_{x_index + 1}'] = None
        x_index += 1
    i_index += 1

add_gear(construction_gear)  # Take into account your gear (for more info, check the function in functions.py)
add_helps()  # Take into account the guildmates' help (for more info, check the function in functions.py)
add_vip(13)  # Take into account your VIP level (for more info, check the function in functions.py)

print(rss_dict)
