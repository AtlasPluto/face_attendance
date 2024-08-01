#referencie 
# https://medium.com/nerd-for-tech/deep-face-recognition-in-python-41522fb47028
# https://viso.ai/computer-vision/deepface/
# https://www.analyticsvidhya.com/blog/2022/04/face-recognition-system-using-python/

import cv2
from PIL import Image, ImageTk
from skuska_class import excel, face_o, gui_c

# nacitanie excel classy

# dolzeite, kamo toto cele prerob na classu lebo toto jhe brutal, uz na tom pracujem excel_f_file uz su classy takze chill oz je tak 45% done (rob dalej lebo to nestihnes),stiho som dorabam poslednu logiku rpe posledne classy a sme chill

########################################################

def kamera():
    
    if gui_c.KAMERA_SLIDER_VALUE > 0 :
        # print(KAMERA_SLIDER_VALUE)
        OK, frame = face_o.video_capture.read()
        face_o.frame = frame
        if not OK :
            return
        if gui_c.DETEKCIA_SLIDER_VALUE > 0 :
            face_o.faces = face_o.detect_bounding_box() 

        else:
            pass
        
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

        photo = Image.fromarray(frame_rgb)
        photo = ImageTk.PhotoImage(image=photo)

        gui_c.label_kamera.configure(image=photo)
        gui_c.label_kamera.image = photo

        gui_c.gui.after(33, kamera)
    else:
        gui_c.label_kamera.configure(image=gui_c.off_img)
        gui_c.label_kamera.image = gui_c.off_img
        
        gui_c.gui.after(33, kamera)

   
kamera()  
gui_c.update_label_date() 
gui_c.gui.mainloop()   

face_o.video_capture.release()    
# ulozenie zmien ktore sme spravili v exceli

# zatvorenie excelu
excel.workbook.close()     

