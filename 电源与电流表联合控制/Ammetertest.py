from Ammeter_Control import *


if __name__ == "__main__":
    # 初始化设备（根据实际情况修改串口和型号）
    pico = KeithleyPicoammeter(
        port="COM3",  # Windows系统
        # port="/dev/ttyUSB0",  # Linux/macOS系统
        baudrate=9600,
        model=InstrumentModel.MODEL_6487
    )

    try:
        # 1. 建立连接
        if not pico.connect():
            exit(1)

        # 2. 设备自检
        pico.self_test()

        # 3. 电流测量示例（启用中值滤波器）
        pico.set_filter(FilterType.MEDIAN, enable=True, param=3)
        pico.set_auto_range(True)
        current = pico.measure_current()
        if current is not None:
            print(f"当前电流：{current:.9f}A")

        # 4. 6487专属：电阻测量（10V测试电压，2.5mA钳位）
        if pico.model == InstrumentModel.MODEL_6487:
            resistance = pico.measure_resistance(voltage=10.0, clamp_current=2.5e-3)
            if resistance is not None:
                print(f"测量电阻：{resistance:.2f}Ω")

        # 5. 高速数据采集（1000点，2μA量程）
        data = pico.capture_buffer_data(sample_count=1000, range_val=2e-6)
        if data:
            print(f"采集数据示例（前10点）：{data[:10]}")

    except KeyboardInterrupt:
        print("\n用户终止操作")
    finally:
        # 断开连接
        pico.disconnect()