from Ammeter_Control import *


if __name__ == "__main__":
    # 初始化设备（根据实际情况修改串口和型号）
    pico = KeithleyPicoammeter(
        port="COM3",  
        baudrate=9600,
        model=InstrumentModel.MODEL_6485
    )

    try:
        # 1. 建立连接
        if not pico.connect():
            exit(1)

        pico.self_test()

        pico.set_zero_correct()
           
        pico.measure_current()  

    except KeyboardInterrupt:
        print("\n用户终止操作")
    finally:
        # 断开连接
        pico.disconnect()