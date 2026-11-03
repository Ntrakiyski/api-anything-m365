import { ApiError } from "./folder.mjs";

export class OutlookConnection {
  constructor({ verify, refresh, signIn, canRefresh, timeoutMs = 300000 }) {
    this.actions = { verify, refresh, signIn, canRefresh };
    this.timeoutMs = timeoutMs;
    this.checked = false;
    this.pending = null;
    this.controller = null;
    this.state = { status: "disconnected", message: "Connect your Microsoft account to read a folder." };
  }

  get isConnecting() {
    return this.controller !== null;
  }

  snapshot() {
    if (!this.checked && !this.pending) {
      this.checked = true;
      this.state = { status: "checking", message: "Checking your saved Microsoft connection." };
      this.pending = this.checkSaved().finally(() => { this.pending = null; });
    }
    return { ...this.state, canRead: this.state.status === "connected", busy: Boolean(this.pending) };
  }

  async checkSaved() {
    try {
      await this.actions.verify();
      this.connected();
    } catch (error) {
      this.failed(error);
    }
  }

  connected() {
    this.state = { status: "connected", message: "Microsoft is connected. Choose a folder to read." };
  }

  failed(error) {
    if (error instanceof ApiError && error.status === 401) {
      this.state = { status: "disconnected", message: "Your Microsoft connection needs sign-in. Select Connect Microsoft." };
    } else if (error instanceof ApiError && error.code === "AUTH_CANCELLED") {
      this.state = { status: "disconnected", message: "Sign-in cancelled. You can connect again when you are ready." };
    } else {
      if (!(error instanceof ApiError)) console.error("Microsoft connection failed: unexpected error");
      this.state = {
        status: "error",
        message: error instanceof ApiError ? error.message : "Could not connect to Microsoft. Check your connection and try again.",
        error: error instanceof ApiError ? error.code : "CONNECTION_FAILED",
      };
    }
  }

  start() {
    if (this.pending) throw new ApiError("AUTH_BUSY", "A Microsoft connection check is already running.", 409);
    this.checked = true;
    this.controller = new AbortController();
    this.state = { status: "refreshing", message: "Trying your saved login first. No sign-in window unless Microsoft needs one." };
    const controller = this.controller;
    const timer = setTimeout(() => {
      controller.abort(new ApiError("AUTH_TIMEOUT", "Sign-in timed out. Connect again and complete it within five minutes.", 408));
    }, this.timeoutMs);
    timer.unref();
    this.pending = this.connect(controller.signal)
      .catch(error => this.failed(error))
      .finally(() => {
        clearTimeout(timer);
        this.controller = null;
        this.pending = null;
      });
    return this.snapshot();
  }

  async connect(signal) {
    try {
      await this.actions.verify();
      signal.throwIfAborted();
      this.connected();
      return;
    } catch (error) {
      signal.throwIfAborted();
      if (!(error instanceof ApiError) || error.status !== 401) throw error;
    }
    if (this.actions.canRefresh()) {
      try {
        await this.actions.refresh(signal);
        signal.throwIfAborted();
        this.connected();
        return;
      } catch (error) {
        signal.throwIfAborted();
        if (!(error instanceof ApiError) || error.status !== 401) throw error;
      }
    }
    this.state = {
      status: "signing_in",
      message: "Complete Microsoft sign-in in the separate window, then close that window. Its mail loading screen is expected.",
    };
    await this.actions.signIn(signal);
    signal.throwIfAborted();
    this.connected();
  }

  cancel() {
    if (!this.controller) throw new ApiError("NO_SIGN_IN", "No sign-in is in progress.", 409);
    this.state = { status: "cancelling", message: "Closing the Microsoft sign-in window." };
    this.controller.abort(new ApiError("AUTH_CANCELLED", "Sign-in cancelled.", 409));
    return this.snapshot();
  }

  requireLogin() {
    if (!this.pending) {
      this.checked = true;
      this.failed(new ApiError("LOGIN_REQUIRED", "Microsoft sign-in is required.", 401));
    }
  }
}
