// 设备数据配置文件

export interface Device {
	name: string;
	image: string;
	specs: string;
	description: string;
	link: string;
}

// 设备类别类型，支持品牌和自定义类别
export type DeviceCategory = Record<string, Device[]> & {
	自定义?: Device[];
};

export const devicesData: DeviceCategory = {
	Apple: [
		{
			name: "iPhone 17",
			image: "/images/device/iphone-17.webp",
			specs: "8GB RAM / 256GB Storage",
			description: "A19 Bionic chip, 120Hz ProMotion, all-day battery life.",
			link: "https://www.apple.com/iphone-17/",
		},
	],
	Xiaomi: [
		{
			name: "Xiaomi Pad 5 Pro 12.4",
			image: "/images/device/xiaomi-pad-5-pro-12.4.webp",
			specs: "8GB RAM / 256GB Storage",
			description:
				"Snapdragon 870, 12.4 120Hz display, 67W turbo charging.",
			link: "https://www.mi.com/xiaomipad5pro-124",
		},
	],
};
