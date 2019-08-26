package com.eks.utils;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import lombok.Data;
import lombok.experimental.Accessors;

import java.util.ArrayList;
import java.util.List;

@Data//注解在类上;提供类所有非static且非final属性的get和set方法,final属性只提供get方法,此外还提供了equals、canEqual、hashCode、toString 方法
@Accessors(chain = true)//chain=boolean值，默认false。如果设置为true，setter返回的是此对象，方便链式调用方法
public class MicrosoftTextToSpeechUtils {
    static {
        ComThread.InitSTA();
    }
    private static Dispatch voiceDispatch = new ActiveXComponent("Sapi.SpVoice").getObject();//声音对象
    private static Dispatch spFileStreamDispatch = new ActiveXComponent("Sapi.SpFileStream").getObject();//音频文件输出流对象，在读取或保存音频文件时使用
    private static Dispatch spMMAudioOutDispatch = new ActiveXComponent("Sapi.SpMMAudioOut").getObject();//音频输出对象
    private static Dispatch sapiSpAudioFormatDispatch = new ActiveXComponent("Sapi.SpAudioFormat").getObject();//音频格式对象

    private static Integer rateInteger = 0;// 频率: -10到10
    private static Integer volumeInteger = 100;// 声音: 1到100
    private static Integer formatTypeInteger = 22;// 音频的输出格式，默认为：SAFT22kHz16BitMono

    private static Integer voiceInteger = 0;// 语音库序号
    private static Integer audioInteger = 0;// 输出设备序号
    //暂停播放语音
    public static void pause(Dispatch voiceDispatch){
        Dispatch.call(voiceDispatch,"Pause");
    }
    //停止播放语音
    public static void stop(Dispatch voiceDispatch){
        Dispatch.call(voiceDispatch,"Stop");
    }
    //改变语音库
    public static void changeVoice(Dispatch voiceDispatch,int voiceInt){
        Dispatch dispatch = Dispatch.call(voiceDispatch,"GetVoices").toDispatch();
        int count = Integer.valueOf(Dispatch.call(dispatch,"Count").toString());
        if(count > 0){
            dispatch = Dispatch.call(dispatch,"Item",new Variant(voiceInt)).toDispatch();
            Dispatch.put(voiceDispatch,"Voice",dispatch);
        }
    }
    //改变音频输出设备
    public static void changeAudioOutput(Dispatch voiceDispatch,int audioInt){
        Dispatch dispatch = Dispatch.call(voiceDispatch,"GetAudioOutputs").toDispatch();
        int count = Integer.valueOf(Dispatch.call(dispatch,"Count").toString());
        if(count > 0){
            Dispatch audioOutput = Dispatch.call(dispatch,"Item",new Variant(audioInt)).toDispatch();
            Dispatch.put(voiceDispatch,"AudioOutput",audioOutput);
        }
    }
    //获取系统中所有的语音库名称
    public static List<String> getVoiceNameStringList(Dispatch voiceDispatch){
        List<String> voiceNameStringList = new ArrayList<>();
        Dispatch dispatch = Dispatch.call(voiceDispatch,"GetVoices").toDispatch();
        int countInt = Integer.valueOf(Dispatch.call(dispatch,"Count").toString());
        if(countInt > 0){
            for(int i = 0;i < countInt;i++){
                Dispatch itemDispatch = Dispatch.call(dispatch,"Item",new Variant(i)).toDispatch();
                voiceNameStringList.add(Dispatch.call(itemDispatch,"GetDescription").toString());
            }
        }
        return voiceNameStringList;
    }
    //获取音频输出设备名称
    public static List<String> getAudioOutputNameStringList(Dispatch voiceDispatch){
        List<String> audioOutputNameStringList = new ArrayList<>();
        Dispatch dispatch = Dispatch.call(voiceDispatch,"GetAudioOutputs").toDispatch();
        int countInt = Integer.valueOf(Dispatch.call(dispatch,"Count").toString());
        if(countInt > 0){
            for(int i = 0;i < countInt;i++){
                Dispatch voiceItem = Dispatch.call(dispatch,"Item",new Variant(i)).toDispatch();
                audioOutputNameStringList.add(Dispatch.call(voiceItem,"GetDescription").toString());
            }
        }
        return audioOutputNameStringList;
    }
    public static void speakBaseProjectPath(String relativePathString) throws Exception {
        String contentString = FileUtils.getContentStringBaseProjectPath(relativePathString);
        speak(contentString,formatTypeInteger,volumeInteger,rateInteger);
    }
    //播放语音
    public static void speak(String contentString,Integer formatTypeInteger,Integer volumeInteger,Integer rateInteger){
        Dispatch.put(sapiSpAudioFormatDispatch,"Type",new Variant(formatTypeInteger));
        //调整音量和读的速度
        Dispatch.put(voiceDispatch,"Volume",new Variant(volumeInteger));// 设置音量
        Dispatch.put(voiceDispatch,"Rate",new Variant(rateInteger));// 设置速率
        Dispatch.putRef(spMMAudioOutDispatch,"Format", sapiSpAudioFormatDispatch);
        Dispatch.put(voiceDispatch,"AllowAudioOutputFormatChangesOnNextSet",new Variant(false));
        Dispatch.putRef(voiceDispatch,"AudioOutputStream", spMMAudioOutDispatch);
        //开始朗读
        Dispatch.call(voiceDispatch,"Speak",new Variant(contentString));
    }
    public static void saveToWav(String relativePathString) throws Exception {
        String contentString = FileUtils.getContentStringBaseProjectPath(relativePathString);
        String[] relativePathStringArray = relativePathString.split("\\.");
        relativePathString = relativePathString.substring(0,relativePathString.length() - relativePathStringArray[relativePathStringArray.length - 1].length() - 1) + ".wav";
        MicrosoftTextToSpeechUtils.saveToWav(contentString,FileUtils.generatePathBaseProjectPath(relativePathString));
    }
    //将文字转换成音频信号，然后输出到.WAV文件
    public static void saveToWav(String contentString,String voiceFilePathString){
        //设置音频流格式类型
        Dispatch.put(sapiSpAudioFormatDispatch,"Type",new Variant(formatTypeInteger));
        //设置文件输出流的格式
        Dispatch.putRef(spFileStreamDispatch,"Format", sapiSpAudioFormatDispatch);
        //调用输出文件流对象的打开方法，创建一个.wav文件
        Dispatch.call(spFileStreamDispatch,"Open",new Variant(voiceFilePathString),new Variant(3),new Variant(true));
        //设置声音对象的音频输出流为输出文件流对象
        Dispatch.putRef(voiceDispatch,"AudioOutputStream", spFileStreamDispatch);
        //调整音量和读的速度
        Dispatch.put(voiceDispatch,"Volume",new Variant(volumeInteger));//设置音量
        Dispatch.put(voiceDispatch,"Rate",new Variant(rateInteger));//设置速率
        //开始朗读
        Dispatch.call(voiceDispatch,"Speak",new Variant(contentString));
        //关闭输出文件流对象，释放资源
        Dispatch.call(spFileStreamDispatch,"Close");
        Dispatch.putRef(voiceDispatch,"AudioOutputStream",null);
    }
}