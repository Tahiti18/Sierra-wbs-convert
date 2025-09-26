#!/usr/bin/env python3

from wbs_ordered_converter import WBSOrderedConverter
import pandas as pd

def analyze_name_mapping():
    """Analyze and fix the name mapping between Sierra and WBS"""
    
    print("NAME MAPPING ANALYSIS")
    print("=" * 80)
    
    # Initialize converter
    converter = WBSOrderedConverter()
    
    # Get Sierra employee names
    employee_hours = converter.parse_sierra_file("/home/user/webapp/Sierra_Payroll_9_12_25_for_Marwan_2.xlsx")
    sierra_names = set(employee_hours.keys())
    
    print(f"Sierra employees found: {len(sierra_names)}")
    print("Sierra names:")
    for name in sorted(sierra_names):
        print(f"  '{name}'")
    
    # Get WBS master names
    wbs_names = set(converter.wbs_order)
    
    print(f"\nWBS master employees: {len(wbs_names)}")
    print("WBS names (first 20):")
    for name in list(converter.wbs_order)[:20]:
        print(f"  '{name}'")
    
    # Check which Sierra names match WBS names exactly
    exact_matches = sierra_names & wbs_names
    print(f"\nExact matches: {len(exact_matches)}")
    for name in sorted(exact_matches):
        print(f"  ✅ '{name}'")
    
    # Find Sierra names that don't match
    no_matches = sierra_names - wbs_names
    print(f"\nSierra names with no exact match: {len(no_matches)}")
    for name in sorted(no_matches):
        print(f"  ❌ '{name}'")
        
        # Try to find partial matches
        possible_matches = []
        for wbs_name in wbs_names:
            # Check if last name matches
            sierra_last = name.split(',')[0].strip() if ',' in name else name.split()[-1]
            wbs_last = wbs_name.split(',')[0].strip() if ',' in wbs_name else wbs_name.split()[-1]
            
            if sierra_last.lower() == wbs_last.lower():
                possible_matches.append(wbs_name)
        
        if possible_matches:
            print(f"     Possible matches: {possible_matches}")

def create_name_mapping_dict():
    """Create a mapping dictionary for Sierra -> WBS names"""
    
    print("\n" + "=" * 80)
    print("CREATING NAME MAPPING")
    print("=" * 80)
    
    # Manual mapping for the cases that need it
    name_mappings = {
        # Based on the analysis, create mappings
        "Martinez, Alberto": "Martinez, Alberto",  # Should match if exists
        "Martinez, Alberto O.": "Olivares, Alberto M",  # Might be this person
        "Gonzalez, Alejandro": "Gonzalez, Alejandro",
        "Lopez, Alexander": "Lopez, Gerwin A",  # Might be this person
        "Castaneda, Andy": "Castaneda, Andy",
        "Rodriguez, Anthony": "Rodriguez, Antoni",  # Note: Antoni vs Anthony
        "Torres, Anthony": "Torres, Anthony",
        "Contreras, Brian": "Contreras, Brian", 
        "Cuevas, Carlos": "Cuevas Barragan, Carlos",  # Might be this person
        "Hernandez, Carlos": "Cuevas Barragan, Carlos",  # Or this might be different
        "Padilla, Carlos": "Padilla, Carlos",
        "Zamora, Cesar": "Zamora, Cesar",  # Need to check if exists in WBS
        "Carrasco, Daniel": "Lopez, Daniel",  # Might be this person
        "Lopez, Daniel": "Lopez, Daniel",
        "Chavez, Derick": "Chavez, Derick J",
        "Robleza, Dianne": "Robleza, Dianne",  # ✅ Exact match
        "Hernandez, Diego": "Hernandez, Diego",
        "Perez, Edgar": "Perez, Edgar",
        "Garcia, Eduardo": "Garcia Garcia, Eduardo",
        "Moreno, Eduardo": "Moreno, Eduardo",
        "Hernandez, Edy": "Chavez, Endhy",  # Might be similar
        "Santos, Efrain": "Santos, Efrain",
        "Gonzalez, Emanuel": "Gonzalez, Emanuel",
        "Martinez, Emiliano": "Martinez, Emiliano B",
        "Ramon, Endhy": "Chavez, Endhy",
        "Vera, Erick": "Serrano, Erick V",  # Might be this person
        "Duarte, Esau": "Duarte, Esau",
        "Arizmendi, Fernando": "Arizmendi, Fernando",
        "Robledo, Francisco": "Robledo, Francisco",
        "Cardoso, Hipolito": "Cardoso, Hipolito",  # Need to add to WBS if missing
        "Dean, Jake": "Dean, Jacob P",
        "Santos, Javier": "Santos, Javier",
        "Pacheco, Jesus": "Pacheco Estrada, Jesus",
        "Alvarez, Jose (Luis)": "Alvarez, Jose",
        "Bocanegra, Jose": "Bocanegra, Jose",
        "Espinoza, Jose": "Espinoza, Jose Federico",
        "Gomez, Jose": "Gomez, Jose",
        "Torrez, Jose": "Torrez, Jose R",
        "Nava, Juan": "Nava, Juan M",
        "Solis, Juan Romero": "Romero Solis, Juan",  # Name order different
        "Vargas, Karina": "Vargas Pineda, Karina",
        "Cortez, Kevin": "Duarte, Kevin",  # Might be this person
        "Duarte, Kevin": "Duarte, Kevin",
        "Esquivel, Kleber": "Esquivel, Kleber",
        "Alcaraz, Luis": "Alcaraz, Luis",
        "Arroyo, Luis": "Arroyo, Jose",  # Might be this person
        "Bello, Luis": "Bello, Luis",
        "Martinez, Maciel": "Martinez, Maciel",
        "Rivas, Manuel": "Rivas Beltran, Angel M",  # Might be this person
        "Cuevas, Marcelo": "Cuevas, Marcelo",
        "Navichoque, Marlon": "Navichoque, Marlon",  # Need to add if missing
        "Garcia, Miguel": "Garcia, Miguel A",
        "Gonzalez, Miguel": "Gonzalez, Miguel",
        "Pelagio, Miguel": "Pelagio, Miguel Angel",
        "Castillo, Moises": "Castillo, Moises",
        "Ramos, Omar": "Ramos Grana, Omar",
        "Pajarito, Ramon": "Pajarito, Ramon",
        "Gomez, Randel": "Gomez, Randel",  # Need to add if missing
        "Flores, Saul": "Flores, Saul Daniel L",
        "Hernandez, Sergio": "Hernandez, Sergio",
        "Stokes, Symone": "Stokes, Symone",  # ✅ Exact match
        "Valle, Victor": "Valle, Victor",
        "Vera, Victor": "Vera, Victor",
        "Lopez, Yair": "Lopez, Yair A",
        "Lopez, Zeferino": "Lopez, Zeferino"
    }
    
    print("Created mapping for Sierra -> WBS names")
    print(f"Total mappings: {len(name_mappings)}")
    
    return name_mappings

if __name__ == "__main__":
    analyze_name_mapping()
    mapping = create_name_mapping_dict()