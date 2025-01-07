import cmds_for_exel
import time

exel_premia = cmds_for_exel.ExelWB()

exel_premia.get_data()
exel_premia.save_data()

time.sleep(4)

exel_premia.show()

print("Пока!")

time.sleep(1)

