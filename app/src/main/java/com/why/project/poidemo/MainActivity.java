package com.why.project.poidemo;

import android.content.Context;
import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.view.View;

import com.why.project.poilib.PoiUtils;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

public class MainActivity extends AppCompatActivity {

	private Context mContext;

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);

		mContext = this;

		//利用doc模板生成doc文件
		findViewById(R.id.btn_poi_doc).setOnClickListener(new View.OnClickListener() {
			@Override
			public void onClick(View view) {
				try {
					InputStream templetDocStream = getAssets().open("请假单模板2.doc");

					String targetDocPath = mContext.getExternalFilesDir("poi").getPath() + File.separator + "请假单2.doc";//这个目录，不需要申请存储权限

					Map<String, String> dataMap = new HashMap<String, String>();
					dataMap.put("$writeDate$", "2018年10月14日");
					dataMap.put("$name$", "HaiyuKing");
					dataMap.put("$dept$", "移动开发组");
					dataMap.put("$leaveType$", "☑倒休 √年假 ✔事假 ☐病假 ☐婚假 ☐产假 ☐其他");
					dataMap.put("$leaveReason$", "倒休一天。");
					dataMap.put("$leaveStartDate$", "2018年10月14日上午");
					dataMap.put("$leaveEndDate$", "2018年10月14日下午");
					dataMap.put("$leaveDay$", "1");
					dataMap.put("$leaveLeader$", "同意");
					dataMap.put("$leaveDeptLeaderImg$", "同意！");

					PoiUtils.writeToDoc(templetDocStream,targetDocPath,dataMap);

				} catch (IOException e) {
					e.printStackTrace();
				}

			}
		});
	}
}
