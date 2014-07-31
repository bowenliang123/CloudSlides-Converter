# -*- coding: UTF-8 -*-

def gen_base_path():
    return "C:\\ppt"

def gen_ppt_path(ppt_id):
    base_dir_path = gen_base_path()
    return base_dir_path + "\\" + ppt_id;

def gen_save_dir_path(ppt_id):
    base_dir_path = gen_base_path()
    return base_dir_path+"\\converted_"+ppt_id;

def gen_single_png_path(ppt_id, index):
    save_dir_path = gen_save_dir_path(ppt_id);
    return save_dir_path + "\\幻灯片" + str(index) + ".JPG";