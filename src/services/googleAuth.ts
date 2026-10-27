import { initializeApp, getApps, getApp } from 'firebase/app';
import { getAuth, signInWithPopup, GoogleAuthProvider, signOut } from 'firebase/auth';
import firebaseConfig from '../../firebase-applet-config.json';

const app = getApps().length > 0 ? getApp() : initializeApp(firebaseConfig);
export const auth = getAuth(app);

export const SCOPES = [
  'email',
  'profile'
];

const provider = new GoogleAuthProvider();
// Chỉ yêu cầu thông tin email và hồ sơ công khai cơ bản để mọi tài khoản Gmail đều đăng nhập được không bị Google chặn
provider.addScope('email');
provider.addScope('profile');

// Luôn hiển thị hộp thoại chọn tài khoản Google để người dùng đăng nhập đúng tài khoản Gmail cá nhân
provider.setCustomParameters({
  prompt: 'select_account'
});

export interface GoogleAuthResult {
  email: string;
  fullName: string;
  photoUrl: string;
  accessToken: string;
}

/**
 * Mở cửa sổ đăng nhập chính thức của Google / Gmail cá nhân (accounts.google.com)
 * Xác thực tài khoản và mật khẩu trực tiếp qua Google
 */
export const signInWithGoogleAccount = async (): Promise<GoogleAuthResult> => {
  const result = await signInWithPopup(auth, provider);
  const credential = GoogleAuthProvider.credentialFromResult(result);
  const accessToken = credential?.accessToken || '';
  const user = result.user;

  if (!user.email) {
    throw new Error('Không nhận được địa chỉ email từ tài khoản Google.');
  }

  return {
    email: user.email,
    fullName: user.displayName || user.email.split('@')[0],
    photoUrl: user.photoURL || '',
    accessToken
  };
};

export const signOutGoogle = async () => {
  try {
    await signOut(auth);
  } catch (e) {
    console.warn('Lỗi khi đăng xuất khỏi tài khoản Google:', e);
  }
};
