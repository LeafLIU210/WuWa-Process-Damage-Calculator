import pandas as pd


class DamageCalculation:
    def __init__(self, atkbs_cr, atkbs_wp, atkpc, atknb, skrmp, dmgif_fc, dmgad_fc, ctrat, ctdmg, lvlcr, lvlmt, resis):
        self.atkbs_cr = atkbs_cr
        self.atkbs_wp = atkbs_wp
        self.atkpc = atkpc
        self.atknb = atknb
        self.skrmp = skrmp
        self.dmgif_fc = dmgif_fc  # List of dmgif factors
        self.dmgad_fc = dmgad_fc  # List of dmgad factors
        self.ctrat = ctrat
        self.ctdmg = ctdmg
        self.lvlcr = lvlcr
        self.lvlmt = lvlmt
        self.resis = resis

    def calculate_dmgifmp(self):
        """
        Calculate the Damage Infliction Multiplier as the sum of factors.
        """
        return sum(self.dmgif_fc)

    def calculate_dmgadmp(self):
        """
        Calculate the Damage Addition Multiplier as the sum of factors.
        """
        return sum(self.dmgad_fc)

    def calculate_dmg(self):
        atkbs = self.atkbs_cr + self.atkbs_wp
        atkmp = atkbs + (1 + self.atkpc / 100) + self.atknb
        ctsmp = self.ctrat * (self.ctdmg - 1) + 1
        dmgifmp = self.calculate_dmgifmp()
        dmgadmp = self.calculate_dmgadmp()
        lvlmp = (self.lvlcr + 100) / (self.lvlmt + self.lvlcr + 200)
        resismp = 1 - self.resis / 100
        dmg = atkmp * self.skrmp * dmgifmp * dmgadmp * ctsmp * lvlmp * resismp
        return dmg


"""
Normal, Heavy, Resonance, Liberation, Cold, Fire, Wind, Electricity, Light, Dark, Else
"""


def read_excel_to_dict(filename):
    df = pd.read_excel(filename, engine='openpyxl')
    params = {}
    # Initialize empty lists for the factors
    dmgad_fc = []
    dmgif_fc = []

    # Iterate over the rows of the DataFrame
    for _, row in df.iterrows():
        param = row['Parameter']
        if param == 'dmgif_fc':
            # Collect all dmgif factors from this row, ignoring NaN values
            dmgif_fc = [row[f'dmgif_{suffix}'] for suffix in
                        ['nm', 'hv', 'rs', 'lb', 'cd', 'fr', 'wd', 'et', 'lt', 'dk', 'el'
                         ] if pd.notna(row[f'dmgif_{suffix}'])]
        elif param == 'dmgad_fc':
            # Collect all dmgad factors from this row, ignoring NaN values
            dmgad_fc = [row[f'dmgad_{suffix}'] for suffix in
                        ['nm', 'hv', 'rs', 'lb', 'cd', 'fr', 'wd', 'et', 'lt', 'dk', 'el'
                         ] if pd.notna(row[f'dmgad_{suffix}'])]
        else:
            # Directly assign the value for parameters that are not lists
            params[param] = row['value']

    # After collecting all factors, add them to the params dictionary
    params['dmgad_fc'] = dmgad_fc
    params['dmgif_fc'] = dmgif_fc
    return params


# Assuming your Excel file is named 'parameters.xlsx'
file_path = 'parameters.xlsx'
params = read_excel_to_dict(file_path)

# Create an instance of DamageCalculation using the unpacked dictionary as keyword arguments
damage_calculator = DamageCalculation(**params)

# Calculate total damage
total_dmg = damage_calculator.calculate_dmg()
print(f"Total Damage: {total_dmg}")

"""
Parameter	value	dmgif_nm	dmgif_hv	dmgif_rs	dmgif_lb	dmgif_cd	dmgif_fr	dmgif_wd	dmgif_et	dmgif_lt	dmgif_dk	dmgif_el	dmgad_nm	dmgad_hv	dmgad_rs	dmgad_lb	dmgad_cd	dmgad_fr	dmgad_wd	dmgad_et	dmgad_lt	dmgad_dk	dmgad_el
atkbs_cr	100																						
atkbs_wp	50																						
atkpc	20																						
atknb	30																						
skrmp	1.5																						
dmgif_fc		1.1 	0.1 	0.1 	0.1 	0.1 	0.1 	0.1 	0.1 	0.1 	0.1 	0.1 											
dmgad_fc													1.1 	0.1 	0.1 	0.1 	0.1 	0.1 	0.1 	0.1 	0.1 	0.1 	0.1 
ctrat	0.2																						
ctdmg	2																						
lvlcr	80																						
lvlmt	70																						
resis	10																						


中文名称Chinese Name	英文名称English Name	变量名variable
总伤害	Total Damage	dmg

攻击力乘区	Attack Power Multiplier	atkmp
基础攻击力	Basic Attack Power	atkbs
角色自身攻击力	Character's Own Attack Power	atkbs_cr
武器基础攻击力	Weapon Basic Attack Power	atkbs_wp
百分比攻击加成	Percentage Attack Bonus	atkpc
普攻百分比攻击加成	Normal Percentage Attack Bonus	atkpc_nm
重击百分比攻击加成	Heavy Percentage Attack Bonus	atkpc_hv
共鸣技能百分比攻击加成	Resonance Percentage Attack Bonus	atkpc_rs
共鸣解放百分比攻击加成	Liberation Percentage Attack Bonus	atkpc_lb
冷凝百分比攻击加成	Cold Percentage Attack Bonus	atkpc_cd
热熔百分比攻击加成	Fire Percentage Attack Bonus	atkpc_fr
气动百分比攻击加成	Wind Percentage Attack Bonus	atkpc_wd
导电百分比攻击加成	Electricity Percentage Attack Bonus	atkpc_et
衍射百分比攻击加成	Light Percentage Attack Bonus	atkpc_lt
湮灭百分比攻击加成	Dark Percentage Attack Bonus	atkpc_dk
数值攻击力	Numeric Attack Power	atknb

技能倍率乘区	Skill Rate Multiplier	skrmp

伤害加深乘区	Damage Intensification Multiplier	dmgifmp
伤害加深百分比	Percentage Damage Intensification	dmgif
普攻伤害加深百分比	Normal Percentage Damage Intensification	dmgif_nm
重击伤害加深百分比	Heavy Percentage Damage Intensification	dmgif_hv
共鸣技能伤害加深百分比	Resonance Percentage Damage Intensification	dmgif_rs
共鸣解放伤害加深百分比	Liberation Percentage Damage Intensification	dmgif_lb
冷凝伤害加深百分比	Cold Percentage Damage Intensification	dmgif_cd
热熔伤害加深百分比	Fire Percentage Damage Intensification	dmgif_fr
气动伤害加深百分比	Wind Percentage Damage Intensification	dmgif_wd
导电伤害加深百分比	Electricity Percentage Damage Intensification	dmgif_et
衍射伤害加深百分比	Light Percentage Damage Intensification	dmgif_lt
湮灭伤害加深百分比	Dark Percentage Damage Intensification	dmgif_dk
全伤害加深百分比	Else Damage Intensification	dmgif_el

伤害加成乘区	Damage Addition Multiplier	dmgadmp
伤害加成百分比	Percentage Damage Addition	dmgad
普攻伤害加成百分比	Normal Percentage Damage Addition	dmgad_nm
重击伤害加成百分比	Heavy Percentage Damage Addition	dmgad_hv
共鸣技能伤害加成百分比	Resonance Percentage Damage Addition	dmgad_rs
共鸣解放伤害加成百分比	Liberation Percentage Damage Addition	dmgad_lb
冷凝伤害加成百分比	Cold Percentage Damage Addition	dmgad_cd
热熔伤害加成百分比	Fire Percentage Damage Addition	dmgad_fr
气动伤害加成百分比	Wind Percentage Damage Addition	dmgad_wd
导电伤害加成百分比	Electricity Percentage Damage Addition	dmgad_et
衍射伤害加成百分比	Light Percentage Damage Addition	dmgad_lt
湮灭伤害加成百分比	Dark Percentage Damage Addition	dmgad_dk
全伤害加成百分比	Else Damage Addition	dmgad_el

暴击乘区	Critical Strike Multiplier	ctsmp
暴击率	Critical Rate	ctrat
暴击伤害	Critical Damage	ctdmg

等级乘区	Level Multiplier	lvlmp
角色等级	Character Level	lvlcr
怪物等级	Monster Level	lvlmt

抗性乘区	Resistance Multiplier	resismp
抗性	Resistance	resis

"""
