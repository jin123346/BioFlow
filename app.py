import sys
from PySide6.QtWidgets import QApplication #PySide6 앱 실행 객체
from ui.main_window import MainWindow #메인 창 클래스 


def main():
    app = QApplication(sys.argv)  #GUI 프로그램 시작 준비
    
    #창 생성하고 화면에 보여줌
    window = MainWindow()
    window.show()

    #이벤트 루프 실행, 버튼 클릭, 창 이동 같은 GUI 동작을 계속 기다림
    sys.exit(app.exec())


if __name__=="__main__":
    main()